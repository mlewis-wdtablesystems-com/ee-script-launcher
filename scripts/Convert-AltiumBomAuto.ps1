<#
.SYNOPSIS
  Convert the latest Altium BOM Excel file for a part number to CSV.

.DESCRIPTION
  No file picker. Given a part number, this script:
    1. Calls Open-AltiumPartFolder.ps1 -Open ResolveJson to resolve the
       Altium folder, latest Rev/Ver subfolder, and production document
       folder for that part.
    2. Looks under "<prodDoc>\BOM" (then "<prodDoc>" and the rev folder
       as fallbacks) for the most recently modified .xlsx / .xls file.
    3. Opens that file in Excel COM, saves sheet 1 as CSV beside the xlsx,
       and runs the same post-processing as Convert_BOM_to_CSV.ps1:
         - validate column C on sheet 1 is fully populated (trailing empty rows
           and MH#/TP#-only designator rows are ignored)
         - validate line 2 contains "-15-" or "-10-" (PCBA)
         - validate line 3 contains "-25-" or "-20-" (PCB)
         - drop line 4 if it contains "Designator"
         - drop BOM lines whose designators are only MH# / TP# (mounting holes, test points)
         - drop blank lines and rows that are only commas / quotes / whitespace
           (so trailing Excel empties are not written to the saved CSV)

.PARAMETER PartNumber
  Internal part number (e.g. 272-15-001). Same number used to find the
  Altium folder.

.EXAMPLE
  .\Convert-AltiumBomAuto.ps1 -PartNumber 272-15-001
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$openScript = Join-Path $scriptsDir 'Open-AltiumPartFolder.ps1'
if (-not (Test-Path -LiteralPath $openScript)) {
  throw "Required helper not found: $openScript"
}

# ---------- step 1: resolve Altium folders ----------

Write-Host "=== Resolving Altium folder for $PartNumber ==="
$resolveOut = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File $openScript -PartNumber $PartNumber -Open ResolveJson 2>&1
if ($LASTEXITCODE -ne 0) {
  throw "Open-AltiumPartFolder.ps1 -Open ResolveJson failed (exit $LASTEXITCODE):`n$($resolveOut -join [Environment]::NewLine)"
}
$joined = ($resolveOut | ForEach-Object { "$_" }) -join "`n"
$rxJson = [regex]'---BEGIN-RESOLVE-JSON---\s*(?<j>\{.*\})\s*---END-RESOLVE-JSON---'
$jm = $rxJson.Match($joined)
if (-not $jm.Success) {
  throw "Could not find resolve JSON markers in output. Raw:`n$joined"
}
$resolved  = $jm.Groups['j'].Value | ConvertFrom-Json
$prodDoc   = $resolved.prodDocDir
$targetDir = $resolved.targetDir
Write-Host ("  Production folder : {0}" -f $prodDoc)
Write-Host ("  Target/rev folder : {0}" -f $targetDir)

# ---------- step 2: find newest xlsx/xls ----------

$searchDirs = @(
  (Join-Path $prodDoc 'BOM'),
  $prodDoc,
  $targetDir
) | Where-Object { Test-Path -LiteralPath $_ -PathType Container }

$candidates = @()
$usedDir = $null
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
Write-Host "=== Altium BOM file ==="
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

# ---------- step 3: validate column C, save as CSV ----------

$excelPath = $pickedXlsx.FullName
$folder    = Split-Path -Parent $excelPath
$baseName  = [System.IO.Path]::GetFileNameWithoutExtension($excelPath)
$csvPath   = Join-Path $folder ($baseName + '.csv')

Write-Host ""
Write-Host "=== Converting to CSV ==="
Write-Host ("  Output CSV    : {0}" -f $csvPath)

function Split-BomCsvDesignatorTokens {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
  return @(($Text -split '[,;]') | ForEach-Object { $_.Trim().Trim('"') } | Where-Object { $_ })
}

function Test-DesignatorsAreOnlyMhOrTp {
  param([string[]]$Tokens)
  if ($Tokens.Count -eq 0) { return $false }
  foreach ($t in $Tokens) {
    if ($t -notmatch '^(?i)MH\d+$' -and $t -notmatch '^(?i)TP\d+$') {
      return $false
    }
  }
  return $true
}

function Find-BomSheetHeaderRow {
  param(
    [Parameter(Mandatory)][object]$Sheet,
    [int]$StartRow,
    [int]$LastRow,
    [int]$StartCol,
    [int]$EndCol
  )
  $scanTo = [Math]::Min($LastRow, $StartRow + 29)
  for ([int]$r = $StartRow; $r -le $scanTo; $r++) {
    $hasQty = $false
    $hasDes = $false
    for ([int]$c = $StartCol; $c -le $EndCol; $c++) {
      $v = ("" + $Sheet.Cells.Item($r, $c).Value2).Trim()
      if ($v -imatch '^(quantity|qty)$') { $hasQty = $true }
      if ($v -match '(?i)designator') { $hasDes = $true }
    }
    if ($hasQty -and $hasDes) { return $r }
  }
  for ([int]$r = $StartRow; $r -le $scanTo; $r++) {
    for ([int]$c = $StartCol; $c -le $EndCol; $c++) {
      $v = ("" + $Sheet.Cells.Item($r, $c).Value2).Trim()
      if ($v -match '(?i)designator') { return $r }
    }
  }
  return -1
}

function Get-BomSheetDesignatorColumn {
  param(
    [Parameter(Mandatory)][object]$Sheet,
    [int]$HeaderRow,
    [int]$StartCol,
    [int]$EndCol
  )
  for ([int]$c = $StartCol; $c -le $EndCol; $c++) {
    $h = ("" + $Sheet.Cells.Item($HeaderRow, $c).Value2).Trim()
    if ($h -imatch 'designator') { return $c }
  }
  return -1
}

function Test-ExcelRowDesignatorsAreOnlyMhOrTp {
  param(
    [Parameter(Mandatory)][object]$Sheet,
    [int]$Row,
    [int]$DesignatorCol
  )
  if ($DesignatorCol -lt 1) { return $false }
  $raw = $Sheet.Cells.Item($Row, $DesignatorCol).Value2
  $text = if ($null -eq $raw) { '' } else { ("" + $raw).Trim() }
  return Test-DesignatorsAreOnlyMhOrTp -Tokens (Split-BomCsvDesignatorTokens -Text $text)
}

function Assert-FirstSheetColumnCComplete {
  param([Parameter(Mandatory)][object]$Sheet)
  [void]$Sheet.Activate()
  $used = $Sheet.UsedRange
  if ($null -eq $used) { throw 'First sheet has no data range.' }
  $startRow = $used.Row
  $lastRowUsed = $startRow + $used.Rows.Count - 1
  $startCol = $used.Column
  $endCol = $used.Column + $used.Columns.Count - 1

  # Excel UsedRange often includes trailing blank rows; ignore rows that are
  # completely empty across all columns in the used rectangle.
  [int]$effectiveLastRow = 0
  for ([int]$r = $lastRowUsed; $r -ge $startRow; $r--) {
    $rowHasContent = $false
    for ([int]$c = $startCol; $c -le $endCol; $c++) {
      $raw = $Sheet.Cells.Item($r, $c).Value2
      if ($null -ne $raw -and -not [string]::IsNullOrWhiteSpace(([string]$raw).Trim())) {
        $rowHasContent = $true
        break
      }
    }
    if ($rowHasContent) {
      $effectiveLastRow = $r
      break
    }
  }
  if ($effectiveLastRow -lt $startRow) {
    throw 'First sheet used range contains only empty rows (no populated cells).'
  }
  if ($effectiveLastRow -lt $lastRowUsed) {
    Write-Host ("  Note: ignoring trailing empty Excel rows {0}-{1}." -f ($effectiveLastRow + 1), $lastRowUsed)
  }

  $headerRow = Find-BomSheetHeaderRow -Sheet $Sheet -StartRow $startRow -LastRow $effectiveLastRow `
    -StartCol $startCol -EndCol $endCol
  $designatorCol = -1
  if ($headerRow -ge 0) {
    $designatorCol = Get-BomSheetDesignatorColumn -Sheet $Sheet -HeaderRow $headerRow `
      -StartCol $startCol -EndCol $endCol
  }

  $skippedMhTp = 0
  for ([int]$r = $startRow; $r -le $effectiveLastRow; $r++) {
    if ($headerRow -ge 0) {
      if ($r -eq $headerRow) { continue }
      if ($r -gt $headerRow -and (Test-ExcelRowDesignatorsAreOnlyMhOrTp -Sheet $Sheet -Row $r -DesignatorCol $designatorCol)) {
        $skippedMhTp++
        continue
      }
    }
    $cell = $Sheet.Cells.Item($r, 3)
    $raw = $cell.Value2
    if ($null -eq $raw -or [string]::IsNullOrWhiteSpace(([string]$raw).Trim())) {
      throw "Column C contains an empty value on Excel row ${r}. Every row in column C within the BOM must be filled."
    }
  }
  if ($skippedMhTp -gt 0) {
    Write-Host ("  Note: column C check skipped for {0} row(s) with only MH# / TP# designators." -f $skippedMhTp)
  }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$workbook = $null

try {
  $workbook = $excel.Workbooks.Open($excelPath, 0, $true)
  if (-not $workbook) { throw "Failed to open workbook: $excelPath" }
  $sheet = $workbook.Worksheets.Item(1)
  Assert-FirstSheetColumnCComplete -Sheet $sheet
  # 6 = xlCSV. Note: SaveAs will overwrite without prompting because
  # DisplayAlerts is off.
  $workbook.SaveAs($csvPath, 6)
  Write-Host "  CSV written."
}
finally {
  if ($workbook) {
    try { $workbook.Close($false) } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
  }
  if ($excel) {
    try { $excel.Quit() } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

# ---------- step 4: wait for file release, then post-process ----------

if (-not (Test-Path -LiteralPath $csvPath)) {
  throw "CSV was not created: $csvPath"
}

$maxRetries = 20
$retry = 0
while ($retry -lt $maxRetries) {
  try {
    $stream = [System.IO.File]::Open($csvPath, 'Open', 'Read', 'None')
    $stream.Close()
    break
  } catch {
    Start-Sleep -Milliseconds 300
    $retry++
  }
}
if ($retry -eq $maxRetries) {
  throw "CSV file still locked after $($maxRetries * 300) ms: $csvPath"
}

Write-Host ""
Write-Host "=== Post-processing CSV ==="
$lines = @(Get-Content -LiteralPath $csvPath)

if ($lines.Count -ge 2 -and $lines[1] -notmatch '-15-' -and $lines[1] -notmatch '-10-') {
  Write-Host "  WARNING: PCBA part number is incorrect (line 2 missing -15- / -10-)."
}
if ($lines.Count -ge 3 -and $lines[2] -notmatch '-25-' -and $lines[2] -notmatch '-20-') {
  Write-Host "  WARNING: PCB part number is incorrect (line 3 missing -25- / -20-)."
}

if ($lines.Count -ge 4 -and $lines[3] -match 'Designator') {
  Write-Host "  Removing line 4 (contains 'Designator')."
  if ($lines.Count -gt 4) {
    $lines = $lines[0..2] + $lines[4..($lines.Count - 1)]
  } else {
    $lines = $lines[0..2]
  }
}

function Find-BomCsvHeaderLineIndex {
  param([string[]]$Lines)
  for ($i = 0; $i -lt $Lines.Count; $i++) {
    $l = $Lines[$i]
    if ($l -match '(?i)Designator' -and $l -match '(?i)Quantity|Qty') { return $i }
  }
  for ($i = 0; $i -lt $Lines.Count; $i++) {
    if ($Lines[$i] -match '(?i)Designator') { return $i }
  }
  return -1
}

function Get-BomCsvDesignatorFieldFromDataLine {
  param(
    [Parameter(Mandatory)][string]$HeaderLine,
    [Parameter(Mandatory)][string]$DataLine
  )
  try {
    $row = @($HeaderLine, $DataLine) | ConvertFrom-Csv | Select-Object -Last 1
    if (-not $row) { return '' }
    $col = @($row.PSObject.Properties.Name | Where-Object { $_ -imatch 'designator' } | Select-Object -First 1)
    if (-not $col) { return '' }
    return ("" + $row.$col).Trim()
  } catch {
    return ''
  }
}

$headerIdx = Find-BomCsvHeaderLineIndex -Lines $lines
$beforeCount = $lines.Count
if ($headerIdx -ge 0) {
  $headerLine = $lines[$headerIdx]
  $kept = New-Object System.Collections.Generic.List[string]
  for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($i -le $headerIdx) {
      [void]$kept.Add($lines[$i])
      continue
    }
    $desField = Get-BomCsvDesignatorFieldFromDataLine -HeaderLine $headerLine -DataLine $lines[$i]
    $tokens = Split-BomCsvDesignatorTokens -Text $desField
    if (Test-DesignatorsAreOnlyMhOrTp -Tokens $tokens) { continue }
    [void]$kept.Add($lines[$i])
  }
  $lines = @($kept)
} else {
  Write-Host '  WARNING: No Designator header row found; skipping MH#/TP# line filter.' -ForegroundColor DarkYellow
}
$dropped = $beforeCount - $lines.Count
if ($dropped -gt 0) {
  Write-Host ("  Dropped {0} line(s) with only MH# / TP# designators." -f $dropped)
}

function Test-IsEffectivelyEmptyCsvLine {
  param([string]$line)
  if ($null -eq $line) { return $true }
  $t = $line.Trim()
  if ($t.Length -eq 0) { return $true }
  # Excel often exports "blank" rows as commas / empty quoted fields only.
  if ($t -match '^[,;\s"]+$') { return $true }
  return $false
}

$beforeEmpty = $lines.Count
$lines = @($lines | Where-Object { -not (Test-IsEffectivelyEmptyCsvLine $_) })
$droppedEmpty = $beforeEmpty - $lines.Count
if ($droppedEmpty -gt 0) {
  Write-Host ("  Dropped {0} empty / separator-only CSV line(s)." -f $droppedEmpty)
}

Set-Content -LiteralPath $csvPath -Value $lines
Write-Host ""
Write-Host ("Done. CSV: {0} ({1} lines)" -f $csvPath, $lines.Count)
exit 0
