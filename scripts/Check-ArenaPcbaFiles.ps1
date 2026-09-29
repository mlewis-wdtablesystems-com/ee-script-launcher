<#
.SYNOPSIS
  Verify PCBA and PCB revisions have the expected files attached and consistent rev letters.

.DESCRIPTION
  Given a PCBA part number PPP-10-XXX (or PPP-15-XXX), the matching PCB
  number is normally PPP-20-XXX (or PPP-25-XXX). If that standard PCB is
  not on the PCBA BOM but another PCB is, that BOM PCB is used instead.
  On the working revision or the most recent checked-in (effective)
  revision of each item this script checks for these file categories:

      PCBA item:
        - Pick and Place
        - Illustrative Image
      PCB (bare board) item:
        - Altium Gerber Files (fabrication zip)

  Also verifies:
    - PCBA revision letter matches PCB revision letter (e.g. A1 vs A).
    - Fabrication zip title revision letter matches the PCBA revision letter.

  The expected revision letter is determined as:
    1. -ExpectedRevision parameter, if given.
    2. The selected revision's `number` (working or most recent checked-in).

.PARAMETER RevisionMode
  working = files and BOM on the working revision (status 0).
  effective_latest = files and BOM on the newest effective (checked-in) revision (status 1).

  Each file's title is checked for a trailing revision segment: "Rev X", "Rev XX",
  "RevX", or "RevXX" (space optional). Mismatches and missing files are flagged.

  When -RevisionMode is effective_latest (most recent checked-in revision), this
  script also runs Compare-WorkingAndCheckedInRevision.ps1 first to confirm the
  working revision has not diverged from the checked-in revision.

  Credentials via env: ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID.

.PARAMETER PartNumber
  PCBA internal part number, e.g. 251-10-001 or 251-15-001.

.PARAMETER ExpectedRevision
  Optional override (e.g. "B" or "B1"). When omitted, the script uses the
  selected revision's number (-RevisionMode).

.PARAMETER SkipPickPlace
  Do not require or verify a Pick and Place file on the PCBA item.

.EXAMPLE
  .\Check-ArenaPcbaFiles.ps1 -PartNumber 251-10-001
.EXAMPLE
  .\Check-ArenaPcbaFiles.ps1 -PartNumber 251-10-001 -SkipPickPlace
  .\Check-ArenaPcbaFiles.ps1 -PartNumber 251-10-001 -RevisionMode effective_latest
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [string]$ExpectedRevision = '',
  [ValidateSet('working', 'effective_latest')] [string]$RevisionMode = 'working',
  [switch]$SkipPickPlace,
  [string]$Email = $env:ARENA_EMAIL,
  [string]$WorkspaceId = $env:ARENA_WORKSPACE_ID,
  [string]$BaseUrl = 'https://api.arenasolutions.com/v1'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $scriptsDir 'AltiumRevisionHelpers.ps1')

function Invoke-ArenaJson {
  param(
    [Parameter(Mandatory = $true)] [string]$Uri,
    [string]$Method = 'Get',
    [hashtable]$Headers,
    [string]$Body,
    [string]$ContentType = 'application/json'
  )
  $params = @{ Uri = $Uri; Method = $Method; UseBasicParsing = $true }
  if ($Headers) { $params.Headers = $Headers }
  if ($Body) { $params.Body = $Body; $params.ContentType = $ContentType }
  $resp = Invoke-WebRequest @params
  $bytes = $resp.RawContentStream.ToArray()
  $text = [System.Text.Encoding]::UTF8.GetString($bytes)
  if ([string]::IsNullOrWhiteSpace($text)) { return $null }
  return $text | ConvertFrom-Json
}

$password = $env:ARENA_PASSWORD
if ([string]::IsNullOrWhiteSpace($Email))       { throw 'Set ARENA_EMAIL or pass -Email.' }
if ([string]::IsNullOrWhiteSpace($password))    { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($WorkspaceId)) { throw 'Set ARENA_WORKSPACE_ID or pass -WorkspaceId.' }

$pnVariants   = Get-BareBoardPartNumberVariants -PartNumber $PartNumber
$pcbaNumber   = $pnVariants.Pcba
$bareNumber   = $pnVariants.Standard
$pcbResolve   = $null

$apiRoot = $BaseUrl.TrimEnd('/')

# Run compare in a separate process before login. A second login from the child
# invalidates the parent's arena_session_id if we authenticate here first.
$compareExit = 0
$compareSummaryLines = New-Object System.Collections.Generic.List[string]
if ($RevisionMode -eq 'effective_latest') {
  $compareScript = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) `
    'Compare-WorkingAndCheckedInRevision.ps1'
  if (-not (Test-Path -LiteralPath $compareScript)) {
    throw "Required script missing: $compareScript"
  }
  Write-Output '=== Compare working vs checked-in revision ==='
  $compareOut = @(
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $compareScript `
      -PartNumber $pcbaNumber -DeferSummary 2>&1 |
      ForEach-Object { "$_" }
  )
  $inCompareSummary = $false
  foreach ($line in $compareOut) {
    if ($line -eq '---BEGIN-COMPARE-SUMMARY---') {
      $inCompareSummary = $true
      continue
    }
    if ($line -eq '---END-COMPARE-SUMMARY---') {
      $inCompareSummary = $false
      continue
    }
    if ($inCompareSummary) {
      [void]$compareSummaryLines.Add($line)
      continue
    }
    Write-Output $line
  }
  $compareExit = $LASTEXITCODE
  if ($null -eq $compareExit) { $compareExit = 0 }
  Write-Output ''
}

$loginBody = @{ email = $Email; password = $password; workspaceId = [long]$WorkspaceId } | ConvertTo-Json
try {
  $login = Invoke-ArenaJson -Uri "$apiRoot/login" -Method Post -Body $loginBody
} catch { throw "Arena login failed: $($_.Exception.Message)" }

$sid = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sid)) { $sid = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sid)) { throw 'Login response did not include arenaSessionId.' }

$headers = @{ arena_session_id = $sid; 'Content-Type' = 'application/json' }

# ---- helpers ----

$revModeLabel = if ($RevisionMode -eq 'working') { 'working' } else { 'most recent checked-in' }

function Select-ArenaRevision {
  param(
    [Parameter(Mandatory = $true)] [object[]]$RevResults,
    [Parameter(Mandatory = $true)] [string]$ModeKey
  )
  if ($ModeKey -eq 'working') {
    $w = $RevResults | Where-Object { $_.status -eq 0 } | Select-Object -First 1
    if (-not $w) { throw "No working revision (status 0) on Arena item." }
    return $w
  }
  $eff = @($RevResults | Where-Object { $_.status -eq 1 })
  if ($eff.Count -eq 0) {
    throw 'No effective (checked-in) revision (status 1). Choose working revision or check in the item first.'
  }
  return $eff | Sort-Object { $_.effectiveDateTime } -Descending | Select-Object -First 1
}

function Resolve-ItemRevisions {
  param(
    [string]$ItemNumber,
    [string]$ModeKey
  )
  $pnEnc = [Uri]::EscapeDataString($ItemNumber)
  $lookup = Invoke-ArenaJson -Uri "$apiRoot/items?number=$pnEnc&limit=1" -Headers $headers
  if (-not $lookup.results -or $lookup.results.Count -eq 0) {
    return @{ Found = $false; ItemNumber = $ItemNumber }
  }
  $anyGuid = $lookup.results[0].guid
  $revs = Invoke-ArenaJson -Uri "$apiRoot/items/$anyGuid/revisions" -Headers $headers
  if (-not $revs.results -or $revs.results.Count -eq 0) {
    return @{ Found = $true; ItemNumber = $ItemNumber; SelectedGuid = $null; SelectedRev = $null }
  }
  $work = $revs.results | Where-Object { $_.status -eq 0 } | Select-Object -First 1
  $effLatest = @($revs.results | Where-Object { $_.status -eq 1 }) |
    Sort-Object { $_.effectiveDateTime } -Descending |
    Select-Object -First 1
  try {
    $selected = Select-ArenaRevision -RevResults @($revs.results) -ModeKey $ModeKey
  } catch {
    return @{
      Found        = $true
      ItemNumber   = $ItemNumber
      SelectedGuid = $null
      SelectedRev  = $null
      Error        = $_.Exception.Message
      WorkingRev   = if ($work) { $work.number } else { $null }
      EffectiveRev = if ($effLatest) { $effLatest.number } else { $null }
    }
  }
  return @{
    Found        = $true
    ItemNumber   = $ItemNumber
    SelectedGuid = $selected.guid
    SelectedRev  = $selected.number
    WorkingRev   = if ($work) { $work.number } else { $null }
    EffectiveRev = if ($effLatest) { $effLatest.number } else { $null }
  }
}

function Get-FilesOnRev {
  param([string]$RevGuid)
  if ([string]::IsNullOrWhiteSpace($RevGuid)) { return @() }
  $resp = Invoke-ArenaJson -Uri "$apiRoot/items/$RevGuid/files" -Headers $headers
  if (-not $resp.results) { return @() }
  return $resp.results
}

function Get-BomOnRev {
  # Returns the immediate child items on a given revision.
  param([string]$RevGuid)
  if ([string]::IsNullOrWhiteSpace($RevGuid)) { return @() }
  $resp = Invoke-ArenaJson -Uri "$apiRoot/items/$RevGuid/bom" -Headers $headers
  if (-not $resp.results) { return @() }
  return $resp.results
}

function Get-FilesByCategoryName {
  param([object[]]$Files, [string]$CategoryName)
  $needle = $CategoryName.Trim().ToLowerInvariant()
  return @($Files | Where-Object {
    $_.file -ne $null -and $_.file.category -ne $null -and (
      ($_.file.category.name -ne $null -and $_.file.category.name.ToString().ToLowerInvariant() -eq $needle) -or
      ($_.file.category.path -ne $null -and $_.file.category.path.ToString().ToLowerInvariant().EndsWith("\$needle"))
    )
  })
}

function Extract-RevFromTitle {
  param([string]$Title)
  if ([string]::IsNullOrWhiteSpace($Title)) { return $null }
  # Accept Rev C, Rev C1, RevC, RevC1, etc. at end of title (space optional).
  $rx = [regex]'(?i)\bRev\s*(?<rev>[A-Za-z0-9]{1,3})\s*$'
  $m = $rx.Match($Title.Trim())
  if ($m.Success) { return $m.Groups['rev'].Value.ToUpperInvariant() }
  return $null
}

function Get-ArenaRevisionLetter {
  param([string]$RevisionNumber)
  if ([string]::IsNullOrWhiteSpace($RevisionNumber)) { return $null }
  try {
    return (Parse-AltiumRevisionToken -RevisionText $RevisionNumber).Letter
  } catch {
    $m = [regex]::Match($RevisionNumber.Trim().ToUpperInvariant(), '^([A-Z]+)')
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
  }
}

function Get-ExpectedRev {
  param($Info, [string]$Override)
  if (-not [string]::IsNullOrWhiteSpace($Override)) {
    return @{ Letter = $Override.Trim().ToUpperInvariant(); Source = "explicit -ExpectedRevision" }
  }
  if ($Info.SelectedRev) {
    $src = if ($RevisionMode -eq 'working') { 'working rev .number' } else { 'checked-in rev .number' }
    return @{ Letter = ($Info.SelectedRev.ToString().ToUpperInvariant()); Source = $src }
  }
  return @{ Letter = $null; Source = 'no rev letter available' }
}

# ---- run checks ----

$pcba = Resolve-ItemRevisions -ItemNumber $pcbaNumber -ModeKey $RevisionMode
$problems = New-Object System.Collections.Generic.List[string]

# ---- resolve PCB from BOM (standard PPP-20/25, else another PCB on the BOM) ----
Write-Output ("[BOM] {0} should contain a bare PCB" -f $pcbaNumber)
$bom = @()
if (-not $pcba.Found) {
  Write-Output "  SKIP: PCBA item not found, cannot inspect BOM."
} elseif (-not $pcba.SelectedGuid) {
  $detail = if ($pcba.Error) { $pcba.Error } else { "PCBA has no $revModeLabel revision." }
  Write-Output ("  SKIP: {0}" -f $detail)
} else {
  $bom = @(Get-BomOnRev -RevGuid $pcba.SelectedGuid)
  Write-Output ("  PCBA {0}-rev BOM: {1} child item(s)" -f $revModeLabel, $bom.Count)
  $pcbResolve = Resolve-BareBoardPartNumberFromBom -PcbaPartNumber $pcbaNumber -BomLines $bom
  $bareNumber = $pcbResolve.Number
  $needle = $bareNumber.Trim().ToLowerInvariant()
  $hit = $bom | Where-Object {
    (Get-BomLinePartNumber -Line $_).ToLowerInvariant() -eq $needle
  } | Select-Object -First 1
  if ($hit) {
    $qty = if ($hit.quantity -ne $null) { $hit.quantity } else { '?' }
    if ($pcbResolve.Source -eq 'bom-alternate') {
      Write-Output ("  -> Standard PCB {0} not in BOM; using {1} (qty={2}, item.guid={3})." -f `
        $pcbResolve.Standard, $bareNumber, $qty, $hit.item.guid)
    } else {
      Write-Output ("  -> Found bare board {0} in BOM (qty={1}, item.guid={2})." -f `
        $bareNumber, $qty, $hit.item.guid)
    }
    if ($pcbResolve.Note) { Write-Output ("  NOTE: {0}" -f $pcbResolve.Note) }
  } else {
    Write-Output ("  WARN: bare board {0} not found in PCBA's BOM." -f $bareNumber)
    [void]$problems.Add("Bare board $bareNumber is not in the BOM of PCBA $pcbaNumber ($revModeLabel rev).")
  }
}
Write-Output ''

Write-Output ("=== {0} (PCBA) <-> {1} (PCB) ===" -f $pcbaNumber, $bareNumber)
Write-Output ("Revision mode: {0}" -f $revModeLabel)
if ($pcbResolve -and $pcbResolve.Source -eq 'bom-alternate') {
  Write-Output ("PCB mapping: BOM alternate (standard {0} not on BOM)" -f $pcbResolve.Standard)
}
Write-Output ''

$bare = Resolve-ItemRevisions -ItemNumber $bareNumber -ModeKey $RevisionMode
Write-Output ''

# ---- PCBA revision letter must match PCB revision letter ----
Write-Output '[Revision letters] PCBA vs PCB'
if (-not $pcba.Found -or -not $bare.Found) {
  Write-Output '  SKIP: one or both items not found in Arena.'
} elseif (-not $pcba.SelectedRev -or -not $bare.SelectedRev) {
  Write-Output '  SKIP: missing revision on one or both items.'
} else {
  $pcbaLetter = Get-ArenaRevisionLetter -RevisionNumber $pcba.SelectedRev
  $bareLetter = Get-ArenaRevisionLetter -RevisionNumber $bare.SelectedRev
  Write-Output ("  PCBA {0}-rev number: {1} (letter {2})" -f $revModeLabel, $pcba.SelectedRev, `
    ($(if ($pcbaLetter) { $pcbaLetter } else { '?' })))
  Write-Output ("  PCB  {0}-rev number: {1} (letter {2})" -f $revModeLabel, $bare.SelectedRev, `
    ($(if ($bareLetter) { $bareLetter } else { '?' })))
  if (-not $pcbaLetter -or -not $bareLetter) {
    Write-Output '  WARN: could not parse revision letter on one or both items.'
    [void]$problems.Add("Could not parse revision letter for PCBA $pcbaNumber and/or PCB $bareNumber.")
  } elseif ($pcbaLetter -eq $bareLetter) {
    Write-Output ("  -> Revision letters match ({0})." -f $pcbaLetter)
  } else {
    Write-Output ("  WARN: revision letters differ (PCBA={0}, PCB={1})." -f $pcbaLetter, $bareLetter)
    [void]$problems.Add("PCBA $pcbaNumber revision letter '$pcbaLetter' does not match PCB $bareNumber revision letter '$bareLetter' ($revModeLabel rev).")
  }
}
Write-Output ''

function Test-RevMatch {
  # Compare a file's title rev letter to the expected revision letter.
  #   Mode = 'strict'     -> exact match (e.g. "B3" must equal "B3")
  #   Mode = 'letter_only'-> compare only the leading [A-Z]+ portion
  #                          (e.g. expected "B3" accepts file rev "B" or "B3")
  param([string]$FileRev, [string]$Expected, [string]$Mode)
  if (-not $FileRev -or -not $Expected) { return $false }
  if ($Mode -eq 'letter_only') {
    $rxLetter = [regex]'^([A-Z]+)'
    $expBase  = $rxLetter.Match($Expected).Groups[1].Value
    $fileBase = $rxLetter.Match($FileRev).Groups[1].Value
    return ($expBase -ne '' -and $expBase -eq $fileBase)
  }
  return ($FileRev -eq $Expected)
}

function Report-Item {
  param(
    $Info,
    [string]$Label,
    [object[]]$RequiredCategories,
    [hashtable]$ExpectedRevOverride = $null
  )

  Write-Output ("[$Label] $($Info.ItemNumber)")
  if (-not $Info.Found) {
    Write-Output "  MISSING: item not found in Arena workspace."
    [void]$script:problems.Add("$Label item $($Info.ItemNumber) not found.")
    Write-Output ''
    return
  }
  if (-not $Info.SelectedGuid) {
    $detail = if ($Info.Error) { $Info.Error } else { "no $revModeLabel revision found." }
    Write-Output ("  MISSING: {0} for $($Info.ItemNumber)." -f $detail)
    [void]$script:problems.Add("$Label item $($Info.ItemNumber) has no $revModeLabel revision.")
    Write-Output ''
    return
  }

  if ($ExpectedRevOverride) {
    $expected = $ExpectedRevOverride
  } else {
    $expected = Get-ExpectedRev -Info $Info -Override $ExpectedRevision
  }
  Write-Output ("  Working rev: number={0}  latest checked-in rev: {1}" -f `
    ($(if ($Info.WorkingRev) { $Info.WorkingRev } else { '<none>' })), `
    ($(if ($Info.EffectiveRev) { $Info.EffectiveRev } else { '<none>' })))
  Write-Output ("  Checking {0} rev: number={1}" -f $revModeLabel, `
    ($(if ($Info.SelectedRev) { $Info.SelectedRev } else { '<none>' })))
  Write-Output ("  Expected 'Rev X' in titles: '{0}' (source: {1})" -f $expected.Letter, $expected.Source)

  $files = @(Get-FilesOnRev -RevGuid $Info.SelectedGuid)
  Write-Output ("  Files on {0} rev: {1}" -f $revModeLabel, $files.Count)

  foreach ($cat in $RequiredCategories) {
    $catName = $cat.Name
    $mode    = $cat.MatchMode
    $matches = @(Get-FilesByCategoryName -Files $files -CategoryName $catName)
    if ($matches.Count -eq 0) {
      Write-Output ("    [{0,-30}] MISSING" -f $catName)
      [void]$script:problems.Add("$Label $($Info.ItemNumber): no '$catName' file on $revModeLabel rev.")
      continue
    }
    foreach ($a in $matches) {
      $title = $a.file.title
      $titleRev = Extract-RevFromTitle -Title $title
      $line = "    [{0,-30}] {1}  '{2}'" -f $catName, $a.file.number, $title
      if (-not $expected.Letter) {
        Write-Output ($line + "  (cannot verify rev: no expected letter)")
      } elseif ($titleRev -eq $null) {
        Write-Output ($line + "  WARN: title has no trailing 'Rev X' or 'RevX'")
        [void]$script:problems.Add("$Label $($Info.ItemNumber) [$catName] '$title' has no trailing 'Rev X' / 'RevX'.")
      } elseif (Test-RevMatch -FileRev $titleRev -Expected $expected.Letter -Mode $mode) {
        $note = if ($mode -eq 'letter_only' -and $titleRev -ne $expected.Letter) {
          "  -> Rev $titleRev OK (letter-only match against expected $($expected.Letter))"
        } else {
          "  -> Rev $titleRev OK"
        }
        Write-Output ($line + $note)
      } else {
        Write-Output ($line + "  WARN: title says Rev $titleRev, expected Rev $($expected.Letter)")
        [void]$script:problems.Add("$Label $($Info.ItemNumber) [$catName] '$title' is Rev $titleRev, expected Rev $($expected.Letter).")
      }
    }
  }
  Write-Output ''
}

# Pick and Place: strict (letter + number). Illustrative Image: letter only.
# Fabrication zip (Altium Gerber Files on PCB): letter must match PCBA revision letter.
$pcbaExpected = Get-ExpectedRev -Info $pcba -Override $ExpectedRevision
$pcbaRequired = @(
  @{ Name = 'Illustrative Image'; MatchMode = 'letter_only' }
)
if (-not $SkipPickPlace) {
  $pcbaRequired = @(
    @{ Name = 'Pick and Place'; MatchMode = 'strict' }
  ) + $pcbaRequired
}
$bareRequired = @(
  @{ Name = 'Altium Gerber Files'; MatchMode = 'letter_only' }
)
$fabricationExpected = @{
  Letter = $pcbaExpected.Letter
  Source = "PCBA rev letter ($pcbaNumber; $($pcbaExpected.Source))"
}

Report-Item -Info $pcba -Label 'PCBA' -RequiredCategories $pcbaRequired
Report-Item -Info $bare -Label 'PCB' -RequiredCategories $bareRequired -ExpectedRevOverride $fabricationExpected

Write-Output '--- Summary ---'
if ($RevisionMode -eq 'effective_latest') {
  Write-Output 'Working vs Checked-In Summary:'
  if ($compareSummaryLines.Count -eq 0) {
    Write-Output '  (compare summary not captured)'
  } else {
    foreach ($line in $compareSummaryLines) { Write-Output $line }
  }
  Write-Output ''
}
Write-Output 'File Check Summary:'
$fileExit = 0
if ($problems.Count -eq 0) {
  Write-Output 'All required files present and revisions consistent.'
} else {
  Write-Output ("{0} issue(s):" -f $problems.Count)
  foreach ($p in $problems) { Write-Output "  - $p" }
  $fileExit = 2
}

if ($compareExit -ne 0 -and $compareExit -ne 2) {
  exit $compareExit
}
if ($compareExit -eq 2 -or $fileExit -eq 2) {
  exit 2
}
exit 0
