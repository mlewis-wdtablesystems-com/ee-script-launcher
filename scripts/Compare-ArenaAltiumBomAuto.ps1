<#
.SYNOPSIS
  Compare Arena BOM vs Altium BOM xlsx automatically, by part number.

.DESCRIPTION
  No file paths required. Given just a part number, this script:
    1. Calls Open-AltiumPartFolder.ps1 -Open ResolveJson to discover the
       Altium folder, latest Rev/Ver subfolder, and the production document
       folder for that part.
    2. Looks under "<prodDoc>\BOM" (then "<prodDoc>" and "<targetDir>" as
       fallbacks) for the most recently modified .xlsx / .xls file.
    3. Logs into Arena and fetches the BOM for the part number, writing it
       to a temporary CSV file (using Get-ArenaBom.ps1 -Format csv).
    4. Calls Compare-ArenaAltiumBom.ps1 -ArenaCsv <tmp.csv> -AltiumXlsx <found>.
    5. Cleans up the temp CSV.

  Credentials via env: ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID.

.PARAMETER PartNumber
  Internal part number (e.g. 272-15-001). The same number is used both for
  the Arena BOM lookup and for resolving the Altium folder.

.PARAMETER Sheet
  Forwarded to Compare-ArenaAltiumBom.ps1 as the primary Altium BOM sheet.
  Default: 1. Sheet 2 (01-PJM Part_Number) is always checked in addition for v2.

.EXAMPLE
  .\Compare-ArenaAltiumBomAuto.ps1 -PartNumber 272-15-001
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [Parameter(Mandatory = $false)] [string]$Sheet = '1'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$openScript    = Join-Path $scriptsDir 'Open-AltiumPartFolder.ps1'
$bomScript     = Join-Path $scriptsDir 'Get-ArenaBom.ps1'
$compareScript = Join-Path $scriptsDir 'Compare-ArenaAltiumBom.ps1'
foreach ($p in @($openScript, $bomScript, $compareScript)) {
  if (-not (Test-Path -LiteralPath $p)) { throw "Required helper script not found: $p" }
}

# ---------- step 1: resolve Altium folders for this part number ----------

Write-Host "=== Resolving Altium folder for $PartNumber ==="
$resolveOut = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File $openScript -PartNumber $PartNumber -Open ResolveJson 2>&1
if ($LASTEXITCODE -ne 0) {
  throw "Open-AltiumPartFolder.ps1 -Open ResolveJson failed (exit $LASTEXITCODE):`n$($resolveOut -join [Environment]::NewLine)"
}

# Extract JSON between markers (the script also prints "Resolved design folder:" etc.)
$joined = ($resolveOut | ForEach-Object { "$_" }) -join "`n"
$rxJson = [regex]'---BEGIN-RESOLVE-JSON---\s*(?<j>\{.*\})\s*---END-RESOLVE-JSON---'
$jm = $rxJson.Match($joined)
if (-not $jm.Success) {
  throw "Could not find resolve JSON markers in output. Raw output:`n$joined"
}
$resolved = $jm.Groups['j'].Value | ConvertFrom-Json
$prodDoc  = $resolved.prodDocDir
$targetDir = $resolved.targetDir
Write-Host ("  Production folder : {0}" -f $prodDoc)
Write-Host ("  Target/rev folder : {0}" -f $targetDir)

# ---------- step 2: find newest xlsx/xls in BOM folder ----------

$searchDirs = @(
  (Join-Path $prodDoc 'BOM'),
  $prodDoc,
  $targetDir
) | Where-Object { Test-Path -LiteralPath $_ -PathType Container }

$candidates = @()
foreach ($d in $searchDirs) {
  $found = @(Get-ChildItem -LiteralPath $d -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Extension -ieq '.xlsx' -or $_.Extension -ieq '.xls' })
  if ($found.Count -gt 0) {
    $candidates = $found
    $usedDir = $d
    break
  }
}

if ($candidates.Count -eq 0) {
  throw ("No .xlsx or .xls file found under any of:`n  - " + ($searchDirs -join "`n  - "))
}

$pickedXlsx = $candidates | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
Write-Host ""
Write-Host ("=== Altium BOM file ===")
Write-Host ("  Search folder : {0}" -f $usedDir)
Write-Host ("  Candidates    : {0}" -f $candidates.Count)
Write-Host ("  Picked        : {0}" -f $pickedXlsx.FullName)
Write-Host ("  Last modified : {0}" -f $pickedXlsx.LastWriteTime)
if ($candidates.Count -gt 1) {
  Write-Host "  Other candidates:"
  foreach ($c in ($candidates | Where-Object { $_.FullName -ne $pickedXlsx.FullName } | Sort-Object LastWriteTime -Descending)) {
    Write-Host ("    - {0}  ({1})" -f $c.Name, $c.LastWriteTime)
  }
}

# ---------- step 3: fetch Arena BOM as CSV ----------

$tempCsv = Join-Path $env:TEMP ("arena_bom_" + [guid]::NewGuid().ToString('N') + ".csv")
Write-Host ""
Write-Host ("=== Fetching Arena BOM for $PartNumber ===")

try {
  # Get-ArenaBom.ps1 emits CSV directly to stdout when -Format csv is used.
  $csvLines = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
    -File $bomScript -PartNumber $PartNumber -Format csv 2>&1
  if ($LASTEXITCODE -ne 0) {
    throw "Get-ArenaBom.ps1 failed (exit $LASTEXITCODE):`n$($csvLines -join [Environment]::NewLine)"
  }
  # Get-ArenaBom emits headers like LineNumber, Quantity, PartNumber, ...
  # Compare-ArenaAltiumBom expects: level, item number, quantity, reference designator, item name.
  # Translate by re-parsing rows and writing a compatible CSV.
  $rows = $csvLines -join "`n" | ConvertFrom-Csv
  if (-not $rows -or $rows.Count -eq 0) {
    throw "Arena BOM for $PartNumber returned 0 rows."
  }
  $translated = $rows | ForEach-Object {
    [pscustomobject]@{
      'level'                = '1'
      'item number'          = $_.PartNumber
      'item name'            = $_.Name
      'quantity'             = $_.Quantity
      'reference designator' = $_.RefDes
    }
  }
  $translated | Export-Csv -LiteralPath $tempCsv -NoTypeInformation -Encoding UTF8
  Write-Host ("  Wrote temp CSV: {0} ({1} lines)" -f $tempCsv, $rows.Count)

  # ---------- step 4: run comparison ----------
  Write-Host ""
  Write-Host "=== Running comparison ==="
  $compareArgs = @(
    '-File', $compareScript,
    '-ArenaCsv', $tempCsv,
    '-AltiumXlsx', $pickedXlsx.FullName,
    '-Sheet', $Sheet,
    '-PjmSpecificationSheet', '2',
    '-AssemblyPartNumber', $PartNumber.Trim()
  )
  & powershell.exe -NoProfile -ExecutionPolicy Bypass @compareArgs
  $compareExit = $LASTEXITCODE
  if ($compareExit -ne 0) {
    Write-Host ""
    Write-Host ("Compare-ArenaAltiumBom.ps1 returned exit code {0}." -f $compareExit)
    exit $compareExit
  }
}
finally {
  if (Test-Path -LiteralPath $tempCsv) {
    Remove-Item -LiteralPath $tempCsv -Force -ErrorAction SilentlyContinue
  }
}

exit 0
