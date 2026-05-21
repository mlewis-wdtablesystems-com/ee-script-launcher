<#
.SYNOPSIS
  Compare Pick & Place CSV descriptions against an Arena assembly BOM by designator.

.DESCRIPTION
  Given a PCBA / assembly part number:
    1. Resolves the Altium production document folder and the single .csv in
       the "Pick Place" subfolder (same rules as Open-AltiumPartFolder).
    2. Fetches the Arena assembly BOM for the chosen revision (working or most
       recent released).
    3. For each BOM line, maps reference designators to Arena item description
       (description when set, else item name) and part number.
    4. Compares each non-fiducial P&P row's Description (or Comment when Description
       is empty) to the Arena text using the same spec-aware regex rules as
       Compare-ArenaAltiumBom / BOM export check (Bom-ArenaAltiumDescriptionChecks.ps1).

  Read-only: does not modify the Pick & Place file.

.PARAMETER PartNumber
  Assembly part number (e.g. 272-15-001).

.PARAMETER RevisionMode
  working = BOM on the working revision (status 0).
  effective_latest = BOM on the newest effective (released) revision (status 1).

.EXAMPLE
  .\Check-PickPlaceFileWithArenaBom.ps1 -PartNumber 272-15-001
  .\Check-PickPlaceFileWithArenaBom.ps1 -PartNumber 272-15-001 -RevisionMode effective_latest
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [ValidateSet('working', 'effective_latest')] [string]$RevisionMode = 'working'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$openScript = Join-Path $scriptsDir 'Open-AltiumPartFolder.ps1'
$checksLib  = Join-Path $scriptsDir 'Bom-ArenaAltiumDescriptionChecks.ps1'
foreach ($p in @($openScript, $checksLib)) {
  if (-not (Test-Path -LiteralPath $p)) { throw "Required script or library missing: $p" }
}
. $checksLib

if ([string]::IsNullOrWhiteSpace($env:ARENA_EMAIL)) { throw 'Set ARENA_EMAIL.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_PASSWORD)) { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_WORKSPACE_ID)) {
  throw 'Set ARENA_WORKSPACE_ID (launcher workspace dropdown sets this per run).'
}

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

function Select-ArenaAssemblyRevision {
  param(
    [Parameter(Mandatory = $true)] [object[]]$RevResults,
    [Parameter(Mandatory = $true)] [string]$ModeKey
  )
  if ($ModeKey -eq 'working') {
    $w = $RevResults | Where-Object { $_.status -eq 0 } | Select-Object -First 1
    if (-not $w) { throw 'No working revision (status 0) on this Arena item.' }
    return $w
  }
  $eff = @($RevResults | Where-Object { $_.status -eq 1 })
  if ($eff.Count -eq 0) {
    throw 'No effective (released) revision (status 1). Choose working revision or release the item first.'
  }
  return $eff | Sort-Object { $_.effectiveDateTime } -Descending | Select-Object -First 1
}

function Split-RefDesignators([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return @() }
  return @(($s -split '[,;]') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Test-ArenaBomLineIsNotFitted {
  param(
    [string]$PartNumber,
    [string]$Name,
    [string]$Notes
  )
  $blob = ("$PartNumber $Name $Notes").Trim()
  if (-not $blob) { return $false }
  return [bool]($blob -match '(?i)\b(not\s*fitted|n\s*/\s*f|dnf|do\s+not\s+fit|not\s+installed)\b')
}

function Test-PickPlaceRowIsFiducial {
  param([hashtable]$Row)
  foreach ($k in @('Designator', 'Comment', 'Description', 'Footprint')) {
    $v = $Row[$k]
    if ($v -and ("$v" -match '(?i)fiducial')) { return $true }
  }
  return $false
}

function Find-PickPlaceHeaderLineIndex {
  param([string[]]$Lines)
  for ($i = 0; $i -lt $Lines.Count; $i++) {
    $l = $Lines[$i]
    if ($l -match '(?i)Designator' -and $l -match '(?i)Comment') {
      return $i
    }
    if ($l -match '(?i)^"Designator"' -and $l -match '(?i)Comment') {
      return $i
    }
  }
  return -1
}

function ConvertFrom-PickPlaceCsvRows {
  param(
    [string]$HeaderLine,
    [string[]]$DataLines
  )
  $block = @($HeaderLine) + @($DataLines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
  $text = $block -join "`r`n"
  return @($text | ConvertFrom-Csv)
}

function Get-PickPlaceDescriptionText {
  param([psobject]$Row)
  $desc = ''
  if ($Row.PSObject.Properties.Name -contains 'Description') {
    $desc = ("$($Row.Description)").Trim()
  }
  if (-not $desc -and ($Row.PSObject.Properties.Name -contains 'Comment')) {
    $desc = ("$($Row.Comment)").Trim()
  }
  return $desc
}

# ---------- Resolve Altium Pick Place CSV ----------

Write-Host "=== Resolving Altium folder for $PartNumber ==="
$resolveOut = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File $openScript -PartNumber $PartNumber.Trim() -Open ResolveJson 2>&1
if ($LASTEXITCODE -ne 0) {
  throw "Open-AltiumPartFolder.ps1 -Open ResolveJson failed (exit $LASTEXITCODE):`n$($resolveOut -join [Environment]::NewLine)"
}
$joined = ($resolveOut | ForEach-Object { "$_" }) -join "`n"
$rxJson = [regex]'---BEGIN-RESOLVE-JSON---\s*(?<j>\{.*\})\s*---END-RESOLVE-JSON---'
$jm = $rxJson.Match($joined)
if (-not $jm.Success) { throw "Could not find resolve JSON markers.`n$joined" }
$resolved = $jm.Groups['j'].Value | ConvertFrom-Json
$prodDoc = $resolved.prodDocDir
Write-Host ("  Production folder : {0}" -f $prodDoc)

$pickPlaceDir = Join-Path $prodDoc 'Pick Place'
if (-not (Test-Path -LiteralPath $pickPlaceDir -PathType Container)) {
  throw "Pick Place folder not found at `"$pickPlaceDir`"."
}
$csvFiles = @(Get-ChildItem -LiteralPath $pickPlaceDir -File -ErrorAction SilentlyContinue |
  Where-Object { $_.Extension -ieq '.csv' })
if ($csvFiles.Count -ne 1) {
  throw "Expected exactly one .csv in Pick Place folder; found $($csvFiles.Count)."
}
$csvPath = $csvFiles[0].FullName
Write-Host ("  Pick & Place CSV  : {0}" -f $csvPath)

$utf8 = New-Object System.Text.UTF8Encoding $false
$lines = [System.IO.File]::ReadAllLines($csvPath, $utf8)
$headerIdx = Find-PickPlaceHeaderLineIndex -Lines $lines
if ($headerIdx -lt 0) { throw "No Designator/Comment header row in: $csvPath" }
$headerLine = $lines[$headerIdx]
$dataLines = @()
for ($i = $headerIdx + 1; $i -lt $lines.Count; $i++) {
  $l = $lines[$i]
  if ([string]::IsNullOrWhiteSpace($l)) { continue }
  if ($l.StartsWith('=')) { break }
  $dataLines += $l
}
$ppRows = ConvertFrom-PickPlaceCsvRows -HeaderLine $headerLine -DataLines $dataLines

# ---------- Arena BOM ----------

Write-Host ''
Write-Host '=== Arena login + BOM ==='
$loginBody = @{
  email       = $env:ARENA_EMAIL
  password    = $env:ARENA_PASSWORD
  workspaceId = [long]$env:ARENA_WORKSPACE_ID
} | ConvertTo-Json
$login = Invoke-ArenaJson -Uri 'https://api.arenasolutions.com/v1/login' -Method Post -Body $loginBody
$sessionId = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sessionId)) { $sessionId = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sessionId)) { throw 'Login response did not include arenaSessionId.' }
$arenaHeaders = @{
  arena_session_id = $sessionId
  'Content-Type'   = 'application/json'
}

$pnEnc = [Uri]::EscapeDataString($PartNumber.Trim())
$lookup = Invoke-ArenaJson -Uri "https://api.arenasolutions.com/v1/items?number=$pnEnc&limit=1" -Headers $arenaHeaders
if (-not $lookup.results -or $lookup.results.Count -eq 0) {
  throw "No Arena item found for part number '$($PartNumber.Trim())'."
}
$parentItem = $lookup.results[0]
$parentGuid = $parentItem.guid
$revs = Invoke-ArenaJson -Uri "https://api.arenasolutions.com/v1/items/$parentGuid/revisions" -Headers $arenaHeaders
if (-not $revs.results -or $revs.results.Count -eq 0) {
  throw "No revisions returned for Arena item '$($PartNumber.Trim())'."
}
$assemblyRev = Select-ArenaAssemblyRevision -RevResults @($revs.results) -ModeKey $RevisionMode
$assemblyRevGuid = ("$($assemblyRev.guid)").Trim()
$assemblyRevNum = if ($null -ne $assemblyRev.number) { ("$($assemblyRev.number)").Trim() } else { '' }
$revLabel = if ($RevisionMode -eq 'working') { 'working' } else { 'most recent released' }
Write-Host ("  Item: {0}  ({1})" -f $parentItem.number, $parentItem.name)
Write-Host ("  BOM revision ({0}): {1}  guid={2}" -f $revLabel, $(if ($assemblyRevNum) { $assemblyRevNum } else { '(no number)' }), $assemblyRevGuid)

$bom = Invoke-ArenaJson -Uri "https://api.arenasolutions.com/v1/items/$assemblyRevGuid/bom" -Headers $arenaHeaders
if (-not $bom.results) {
  throw "Arena BOM returned no lines for $($PartNumber.Trim()) on $revLabel revision."
}

$itemDescCache = @{}
function Get-ArenaItemDescriptionText {
  param([string]$ItemNumber)
  $key = $ItemNumber.Trim()
  if (-not $key) { return '' }
  if ($itemDescCache.ContainsKey($key)) { return $itemDescCache[$key] }
  $enc = [Uri]::EscapeDataString($key)
  $viewEnc = [Uri]::EscapeDataString('full')
  $items = Invoke-ArenaJson -Uri "https://api.arenasolutions.com/v1/items?number=$enc&limit=1&responseview=$viewEnc" -Headers $arenaHeaders
  $side = ''
  if ($items.results -and $items.results.Count -gt 0) {
    $it = $items.results[0]
    $desc = if ($null -ne $it.description) { ("$($it.description)").Trim() } else { '' }
    $nm = if ($null -ne $it.name) { ("$($it.name)").Trim() } else { '' }
    $side = if ($desc) { $desc } else { $nm }
  }
  $itemDescCache[$key] = $side
  return $side
}

$designatorMap = @{}
$nfDesignators = @{}
foreach ($r in $bom.results) {
  $pn = ("$($r.item.number)").Trim()
  $name = ("$($r.item.name)").Trim()
  $notes = if ($null -ne $r.notes) { ("$($r.notes)").Trim() } else { '' }
  $isNf = Test-ArenaBomLineIsNotFitted -PartNumber $pn -Name $name -Notes $notes
  $desc = Get-ArenaItemDescriptionText -ItemNumber $pn
  if (-not $desc) { $desc = $name }
  foreach ($des in @(Split-RefDesignators ("$($r.refDes)"))) {
    if (-not $des) { continue }
    if ($isNf) {
      $nfDesignators[$des] = $true
      continue
    }
    $designatorMap[$des] = [pscustomobject]@{
      PartNumber  = $pn
      Description = $desc
    }
  }
}
Write-Host ("  BOM lines: {0}  fitted designators: {1}  NF/DNF designators: {2}" -f `
  $bom.results.Count, $designatorMap.Count, $nfDesignators.Count)

# ---------- Compare P&P vs Arena ----------

Write-Host ''
Write-Host '=== Pick & Place vs Arena BOM (description check) ==='

$ok = @()
$mismatch = @()
$noArena = @()
$emptyPp = @()
$fiducial = 0
$nfInPp = @()
$ppDesignatorsSeen = @{}

foreach ($row in $ppRows) {
  $des = ("$($row.Designator)").Trim()
  if (-not $des) { continue }
  $ppDesignatorsSeen[$des] = $true

  $ht = @{}
  foreach ($p in $row.PSObject.Properties) { $ht[$p.Name] = $p.Value }
  if (Test-PickPlaceRowIsFiducial -Row $ht) {
    $fiducial++
    continue
  }

  if ($nfDesignators.ContainsKey($des)) {
    $nfInPp += [pscustomobject]@{ Designator = $des }
    continue
  }

  $ppText = Get-PickPlaceDescriptionText -Row $row
  if (-not $ppText) {
    $emptyPp += [pscustomobject]@{ Designator = $des }
    continue
  }

  if (-not $designatorMap.ContainsKey($des)) {
    $noArena += [pscustomobject]@{
        Designator = $des
        PickPlace  = $ppText
      }
    continue
  }

  $arena = $designatorMap[$des]
  $chk = Test-BomDescriptionsCompatible `
    -ArenaDescription $arena.Description `
    -AltiumDescription $ppText `
    -PartNumber $arena.PartNumber

  if ($chk.Ok) {
    $ok += [pscustomobject]@{
        Designator = $des
        PartNumber = $arena.PartNumber
      }
  }
  else {
    $mismatch += [pscustomobject]@{
        Designator = $des
        PartNumber = $arena.PartNumber
        Arena      = $arena.Description
        PickPlace  = $ppText
        Detail     = $chk.Message
      }
  }
}

$arenaOnly = @($designatorMap.Keys | Where-Object { -not $ppDesignatorsSeen.ContainsKey($_) } | Sort-Object)

function Write-ListSection {
  param(
    [string]$Title,
    [System.ConsoleColor]$Color,
    [object[]]$Items,
    [scriptblock]$Formatter
  )
  Write-Host ''
  Write-Host $Title -ForegroundColor $Color
  if (-not $Items -or $Items.Count -eq 0) {
    Write-Host '  (none)' -ForegroundColor DarkGray
    return
  }
  foreach ($item in $Items) {
    & $Formatter $item
  }
}

Write-ListSection -Title "Matched ($($ok.Count))" -Color Green -Items $ok -Formatter {
  param($x)
  Write-Host ("  {0}  {1}" -f $x.Designator, $x.PartNumber) -ForegroundColor DarkGreen
}

Write-ListSection -Title "Description mismatches ($($mismatch.Count))" -Color Red -Items $mismatch -Formatter {
  param($x)
  Write-Host ("  {0}  {1}" -f $x.Designator, $x.PartNumber) -ForegroundColor Red
  Write-Host ("    Pick & Place: {0}" -f $x.PickPlace)
  Write-Host ("    Arena BOM   : {0}" -f $x.Arena)
  if ($x.Detail) { Write-Host ("    Detail    : {0}" -f $x.Detail) -ForegroundColor Yellow }
}

Write-ListSection -Title "P&P designators not on Arena BOM ($($noArena.Count))" -Color Yellow -Items $noArena -Formatter {
  param($x)
  Write-Host ("  {0}  P&P text: {1}" -f $x.Designator, $x.PickPlace) -ForegroundColor Yellow
}

Write-ListSection -Title "Arena BOM designators missing from P&P ($($arenaOnly.Count))" -Color Yellow -Items $arenaOnly -Formatter {
  param($des)
  $a = $designatorMap[$des]
  Write-Host ("  {0}  {1}  Arena: {2}" -f $des, $a.PartNumber, $a.Description) -ForegroundColor Yellow
}

Write-ListSection -Title "P&P rows with empty Description/Comment ($($emptyPp.Count))" -Color DarkYellow -Items $emptyPp -Formatter {
  param($x)
  Write-Host ("  {0}" -f $x.Designator) -ForegroundColor DarkYellow
}

Write-ListSection -Title "P&P rows for Arena NF/DNF lines ($($nfInPp.Count))" -Color DarkGray -Items $nfInPp -Formatter {
  param($x)
  Write-Host ("  {0}  (still present in P&P; Arena marks Not Fitted)" -f $x.Designator) -ForegroundColor DarkGray
}

if ($fiducial -gt 0) {
  Write-Host ''
  Write-Host ("Fiducial rows skipped: {0}" -f $fiducial) -ForegroundColor DarkGray
}

Write-Host ''
$issueCount = $mismatch.Count + $noArena.Count + $arenaOnly.Count + $emptyPp.Count
if ($issueCount -eq 0) {
  Write-Host "=== All compared designators match Arena BOM descriptions ($($ok.Count) checked) ===" -ForegroundColor Green
  exit 0
}

Write-Host ("=== Check finished with issues: {0} mismatch(es), {1} only-in-P&P, {2} only-in-Arena, {3} empty P&P text ===" -f `
  $mismatch.Count, $noArena.Count, $arenaOnly.Count, $emptyPp.Count) -ForegroundColor Red
exit 1
