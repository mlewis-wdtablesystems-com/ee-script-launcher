<#
.SYNOPSIS
  Verify a PCB assembly working revision has the expected files attached
  (PCBA + bare board), and that each file's title ends with the expected
  "Rev X" or "Rev XX" letters.

.DESCRIPTION
  Given a PCBA part number PPP-10-XXX (or PPP-15-XXX), the matching bare
  board number is PPP-20-XXX (or PPP-25-XXX). On the working revision of
  each item this script checks for these file categories:

      PCBA item:
        - Pick and Place
        - Illustrative Image
      Bare board item:
        - Altium Gerber Files

  The expected revision letter is determined as:
    1. -ExpectedRevision parameter, if given.
    2. The working revision's `number` (if Arena has assigned one).
    3. The latest effective revision's `number` as a fallback (with a note).

  Each file's title is checked for a trailing "Rev <X>" or "Rev <XX>"
  segment. Mismatches and missing files are flagged.

  Credentials via env: ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID.

.PARAMETER PartNumber
  PCBA internal part number, e.g. 251-10-001 or 251-15-001.

.PARAMETER ExpectedRevision
  Optional override (e.g. "B" or "B1"). When omitted, the script tries
  the working revision's letter and falls back to the effective letter.

.PARAMETER SkipPickPlace
  Do not require or verify a Pick and Place file on the PCBA item.

.EXAMPLE
  .\Check-ArenaPcbaFiles.ps1 -PartNumber 251-10-001
.EXAMPLE
  .\Check-ArenaPcbaFiles.ps1 -PartNumber 251-10-001 -SkipPickPlace
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [string]$ExpectedRevision = '',
  [switch]$SkipPickPlace,
  [string]$Email = $env:ARENA_EMAIL,
  [string]$WorkspaceId = $env:ARENA_WORKSPACE_ID,
  [string]$BaseUrl = 'https://api.arenasolutions.com/v1'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

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

# Validate + derive bare-board PN.
$m = [regex]::Match($PartNumber.Trim(), '^(?<prefix>\d+)-(?<asm>10|15)-(?<suffix>\d+)$')
if (-not $m.Success) {
  throw "PartNumber must look like PPP-10-XXX or PPP-15-XXX. Got: '$PartNumber'."
}
$asmCode      = $m.Groups['asm'].Value
$bareCode     = if ($asmCode -eq '10') { '20' } else { '25' }
$pcbaNumber   = "$($m.Groups['prefix'].Value)-$asmCode-$($m.Groups['suffix'].Value)"
$bareNumber   = "$($m.Groups['prefix'].Value)-$bareCode-$($m.Groups['suffix'].Value)"

$apiRoot = $BaseUrl.TrimEnd('/')

$loginBody = @{ email = $Email; password = $password; workspaceId = [long]$WorkspaceId } | ConvertTo-Json
try {
  $login = Invoke-ArenaJson -Uri "$apiRoot/login" -Method Post -Body $loginBody
} catch { throw "Arena login failed: $($_.Exception.Message)" }

$sid = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sid)) { $sid = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sid)) { throw 'Login response did not include arenaSessionId.' }

$headers = @{ arena_session_id = $sid; 'Content-Type' = 'application/json' }

# ---- helpers ----

function Resolve-WorkingRev {
  param([string]$ItemNumber)
  $pnEnc = [Uri]::EscapeDataString($ItemNumber)
  $lookup = Invoke-ArenaJson -Uri "$apiRoot/items?number=$pnEnc&limit=1" -Headers $headers
  if (-not $lookup.results -or $lookup.results.Count -eq 0) {
    return @{ Found = $false; ItemNumber = $ItemNumber }
  }
  $anyGuid = $lookup.results[0].guid
  $revs = Invoke-ArenaJson -Uri "$apiRoot/items/$anyGuid/revisions" -Headers $headers
  $work = $revs.results | Where-Object { $_.status -eq 0 } | Select-Object -First 1
  $eff  = $revs.results | Where-Object { $_.status -eq 1 } | Select-Object -First 1
  return @{
    Found        = $true
    ItemNumber   = $ItemNumber
    WorkingGuid  = if ($work) { $work.guid }   else { $null }
    WorkingRev   = if ($work) { $work.number } else { $null }
    EffectiveRev = if ($eff)  { $eff.number }  else { $null }
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
  $rx = [regex]'(?i)\bRev\s+(?<rev>[A-Za-z0-9]{1,3})\s*$'
  $m = $rx.Match($Title.Trim())
  if ($m.Success) { return $m.Groups['rev'].Value.ToUpperInvariant() }
  return $null
}

function Get-ExpectedRev {
  param($Info, [string]$Override)
  if (-not [string]::IsNullOrWhiteSpace($Override)) {
    return @{ Letter = $Override.Trim().ToUpperInvariant(); Source = "explicit -ExpectedRevision" }
  }
  if ($Info.WorkingRev) {
    return @{ Letter = ($Info.WorkingRev.ToString().ToUpperInvariant()); Source = "working rev .number" }
  }
  if ($Info.EffectiveRev) {
    return @{ Letter = ($Info.EffectiveRev.ToString().ToUpperInvariant()); Source = "fallback to effective rev .number (working rev has no number)" }
  }
  return @{ Letter = $null; Source = 'no rev letter available' }
}

# ---- run checks ----

Write-Output ("=== {0} (PCBA) <-> {1} (bare board) ===" -f $pcbaNumber, $bareNumber)
Write-Output ''

$pcba = Resolve-WorkingRev -ItemNumber $pcbaNumber
$bare = Resolve-WorkingRev -ItemNumber $bareNumber

$problems = New-Object System.Collections.Generic.List[string]

# ---- bare board must appear in the PCBA's BOM ----
Write-Output ("[BOM] {0} should contain {1}" -f $pcbaNumber, $bareNumber)
if (-not $pcba.Found) {
  Write-Output "  SKIP: PCBA item not found, cannot inspect BOM."
} elseif (-not $pcba.WorkingGuid) {
  Write-Output "  SKIP: PCBA has no working revision, cannot inspect BOM."
} else {
  $bom = @(Get-BomOnRev -RevGuid $pcba.WorkingGuid)
  Write-Output ("  PCBA working-rev BOM: {0} child item(s)" -f $bom.Count)
  $needle = $bareNumber.Trim().ToLowerInvariant()
  $hit = $bom | Where-Object {
    $_.item -ne $null -and $_.item.number -ne $null -and
    $_.item.number.ToString().Trim().ToLowerInvariant() -eq $needle
  } | Select-Object -First 1
  if ($hit) {
    $qty = if ($hit.quantity -ne $null) { $hit.quantity } else { '?' }
    Write-Output ("  -> Found bare board in BOM (qty={0}, item.guid={1})." -f $qty, $hit.item.guid)
  } else {
    Write-Output "  WARN: bare board not found in PCBA's BOM."
    [void]$problems.Add("Bare board $bareNumber is not in the BOM of PCBA $pcbaNumber working rev.")
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
  param($Info, [string]$Label, [object[]]$RequiredCategories)

  Write-Output ("[$Label] $($Info.ItemNumber)")
  if (-not $Info.Found) {
    Write-Output "  MISSING: item not found in Arena workspace."
    [void]$script:problems.Add("$Label item $($Info.ItemNumber) not found.")
    Write-Output ''
    return
  }
  if (-not $Info.WorkingGuid) {
    Write-Output "  MISSING: no working revision found for $($Info.ItemNumber)."
    [void]$script:problems.Add("$Label item $($Info.ItemNumber) has no working revision.")
    Write-Output ''
    return
  }

  $expected = Get-ExpectedRev -Info $Info -Override $ExpectedRevision
  Write-Output ("  Working rev: number={0}  effective rev: {1}" -f `
    ($(if ($Info.WorkingRev) { $Info.WorkingRev } else { '<none>' })), `
    ($(if ($Info.EffectiveRev) { $Info.EffectiveRev } else { '<none>' })))
  Write-Output ("  Expected 'Rev X' in titles: '{0}' (source: {1})" -f $expected.Letter, $expected.Source)

  $files = @(Get-FilesOnRev -RevGuid $Info.WorkingGuid)
  Write-Output ("  Files on working rev: {0}" -f $files.Count)

  foreach ($cat in $RequiredCategories) {
    $catName = $cat.Name
    $mode    = $cat.MatchMode
    $matches = @(Get-FilesByCategoryName -Files $files -CategoryName $catName)
    if ($matches.Count -eq 0) {
      Write-Output ("    [{0,-30}] MISSING" -f $catName)
      [void]$script:problems.Add("$Label $($Info.ItemNumber): no '$catName' file on working rev.")
      continue
    }
    foreach ($a in $matches) {
      $title = $a.file.title
      $titleRev = Extract-RevFromTitle -Title $title
      $line = "    [{0,-30}] {1}  '{2}'" -f $catName, $a.file.number, $title
      if (-not $expected.Letter) {
        Write-Output ($line + "  (cannot verify rev: no expected letter)")
      } elseif ($titleRev -eq $null) {
        Write-Output ($line + "  WARN: title has no trailing 'Rev <X>'")
        [void]$script:problems.Add("$Label $($Info.ItemNumber) [$catName] '$title' has no trailing 'Rev X'.")
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

# Pick and Place: strict (letter + number). Illustrative Image and Gerber files:
# letter portion only (e.g. expected A1 accepts Rev A on fabrication zip title).
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

Report-Item -Info $pcba -Label 'PCBA'       -RequiredCategories $pcbaRequired
Report-Item -Info $bare -Label 'Bare board' -RequiredCategories $bareRequired

Write-Output '--- Summary ---'
if ($problems.Count -eq 0) {
  Write-Output 'All required files present and revisions consistent.'
  exit 0
} else {
  Write-Output ("{0} issue(s):" -f $problems.Count)
  foreach ($p in $problems) { Write-Output "  - $p" }
  exit 2
}
