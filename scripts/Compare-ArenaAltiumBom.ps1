<#
.SYNOPSIS
  Compare an Arena BOM CSV export against a sheet of an Altium BOM xlsx.

.DESCRIPTION
  Matches line items by internal part number ("item number" in Arena, the
  "*Part_Number" header in Altium). Reports:
    - Parts only in Arena
    - Parts only in Altium
    - Altium rows without a part number (mounting holes, test points, etc.)
    - Common parts whose quantity or reference-designator set differs
    - Common parts: spec-aware Arena item name vs Altium description (voltage,
      current, footprint codes, passive-style values; tolerant of spacing/case)
    - Optional: sheet 2 (default) with "01-PJM Part_Number" and
      "03-Description General" (when present) vs Arena

  Reads xlsx via Excel COM (your machine already has Excel for the Altium
  Convert_BOM_to_CSV.ps1 workflow). The xlsx is copied to TEMP first so a
  file open in Excel does not block the read.

.PARAMETER ArenaCsv
  Path to the Arena BOM CSV. Expects header columns:
    level, item number, item name, quantity, reference designator
  Only level == 1 rows are considered.

.PARAMETER AltiumXlsx
  Path to the Altium BOM xlsx (Project Outputs BOM).

.PARAMETER Sheet
  Sheet name or 1-based index in the Altium xlsx. Default: 2 (the engineering
  BOM sheet, e.g. "ENG BOM"). The header row is auto-detected by looking for
  the row that contains "Quantity" and a "Designator" column.

.PARAMETER PjmSpecificationSheet
  1-based sheet index (or name) for the PJM export tab that contains
  "01-PJM Part_Number" and a description column ("03-Description General"
  preferred; 01-PJM ... Description or Description as fallbacks). Default ''
  (skip). The automatic v2 wrapper passes 2 so sheet 2 is checked when present.

.PARAMETER AssemblyPartNumber
  Top-level assembly being compared (e.g. 153-15-009). That part number is
  omitted from the Altium side of the diff because Arena BOM exports list
  child lines only, not the parent assembly row present in Altium exports.

.EXAMPLE
  .\Compare-ArenaAltiumBom.ps1 -ArenaCsv 'C:\...\292-15-001.csv' `
      -AltiumXlsx 'C:\...\292-15-001-ASY RevA LED Board_20260409.xlsx'
#>
param(
  [Parameter(Mandatory = $true)] [string]$ArenaCsv,
  [Parameter(Mandatory = $true)] [string]$AltiumXlsx,
  [Parameter(Mandatory = $false)] [string]$Sheet = '2',
  [Parameter(Mandatory = $false)] [string]$PjmSpecificationSheet = '',
  [Parameter(Mandatory = $false)] [string]$AssemblyPartNumber = ''
)

$ErrorActionPreference = 'Stop'

. (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) 'Bom-ArenaAltiumDescriptionChecks.ps1')

function Write-Section([string]$title, [System.ConsoleColor]$color = 'Cyan') {
  Write-Host ''
  Write-Host "=== $title ===" -ForegroundColor $color
}

function Split-Designators([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return @() }
  return @(($s -split '[,;]') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Test-IsAltiumBomNonDataPartNumberRow {
  <#
  Rows to omit from the primary Altium BOM diff: column-heading text copied into
  the part-number column, and the top-level assembly line (Arena lists children only).
  #>
  param(
    [Parameter(Mandatory = $true)] [string]$PartNumber,
    [Parameter(Mandatory = $false)] [string]$AssemblyPartNumber = '',
    [int]$Quantity = 0,
    [string[]]$Designators = @()
  )
  $pn = $PartNumber.Trim()
  if (-not $pn) { return $false }

  if ($pn -imatch '(?i)^01-PJM\s+Part[_ ]?Number$') { return $true }
  if ($pn -imatch '(?i)^03-Description(\s+General)?$') { return $true }
  if ($pn -imatch '(?i)^(WDTS\s*)?Part[_ ]?Number$') { return $true }
  if ($pn -imatch '(?i)^ITEM\s*NUMBER$') { return $true }
  if ($pn -imatch '(?i)^Internal\s*Part\s*Number$') { return $true }
  if ($pn -imatch '(?i)^Quantity$' -or $pn -imatch '(?i)^Qty$') { return $true }

  if (-not [string]::IsNullOrWhiteSpace($AssemblyPartNumber)) {
    if ($pn -ieq $AssemblyPartNumber.Trim()) { return $true }
  }

  # Secondary header row: part-number cell is a heading and designator cell is "Designator".
  if ($Quantity -eq 0 -and $Designators.Count -eq 1) {
    if ($Designators[0] -imatch '^(?i)Designator(s)?$') { return $true }
  }

  return $false
}

if (-not (Test-Path -LiteralPath $ArenaCsv)) {
  Write-Host "Arena CSV not found: $ArenaCsv" -ForegroundColor DarkRed
  exit 1
}
if (-not (Test-Path -LiteralPath $AltiumXlsx)) {
  Write-Host "Altium xlsx not found: $AltiumXlsx" -ForegroundColor DarkRed
  exit 1
}

# --------------------------- Load Arena CSV ---------------------------
$arena = @{}
try {
  $arenaRows = Import-Csv -LiteralPath $ArenaCsv
} catch {
  Write-Host "Failed to read Arena CSV: $($_.Exception.Message)" -ForegroundColor DarkRed
  exit 1
}

# Resolve column names tolerantly (handles UTF-8 BOM etc.)
$firstRow = $arenaRows | Select-Object -First 1
if (-not $firstRow) {
  Write-Host "Arena CSV is empty." -ForegroundColor DarkRed
  exit 1
}
$arenaCols = $firstRow.PSObject.Properties.Name
function Find-ArenaCol([string[]]$candidates) {
  foreach ($c in $candidates) {
    foreach ($name in $arenaCols) {
      if ($name.Trim().ToLowerInvariant() -eq $c.ToLowerInvariant()) { return $name }
    }
  }
  return $null
}
$colLevel = Find-ArenaCol @('level')
$colPN    = Find-ArenaCol @('item number','part number','internal part number')
$colName  = Find-ArenaCol @('item name','description')
$colQty   = Find-ArenaCol @('quantity','qty')
$colDes   = Find-ArenaCol @('reference designator','reference designators','designator','designators')

if (-not $colPN -or -not $colQty -or -not $colDes) {
  Write-Host "Arena CSV is missing expected columns. Found: $($arenaCols -join ', ')" -ForegroundColor DarkRed
  exit 1
}

foreach ($row in $arenaRows) {
  if ($colLevel -and ($row.$colLevel) -ne '1') { continue }
  $pn = ("" + $row.$colPN).Trim()
  if (-not $pn) { continue }
  $qty = 0
  [int]::TryParse(("" + $row.$colQty), [ref]$qty) | Out-Null
  $arena[$pn] = [pscustomobject]@{
    Quantity    = $qty
    Description = if ($colName) { ("" + $row.$colName).Trim() } else { '' }
    Designators = Split-Designators ($row.$colDes)
  }
}

# --------------------------- Load Altium xlsx (COM) ---------------------------
$excel = $null
$wb = $null
$tempCopy = $null
$altium = @{}
$noPnRows = @()
$resolvedSheetName = $null
$pjmMismatches = New-Object System.Collections.Generic.List[object]
$pjmDupPnGroups = @()

try {
  try {
    $excel = New-Object -ComObject Excel.Application
  } catch {
    Write-Host "Could not start Excel COM. Make sure Excel is installed." -ForegroundColor DarkRed
    exit 2
  }
  $excel.Visible = $false
  $excel.DisplayAlerts = $false

  # Must keep the real extension (.xls vs .xlsx / .xlsm). Copying a legacy .xls
  # bytes to a path ending in .xlsx makes Excel COM reject the file as invalid.
  $srcExt = [System.IO.Path]::GetExtension($AltiumXlsx)
  if ([string]::IsNullOrWhiteSpace($srcExt)) { $srcExt = '.xlsx' }
  $tempCopy = Join-Path $env:TEMP ("altium_bom_" + [guid]::NewGuid().ToString('N') + $srcExt)
  Copy-Item -LiteralPath $AltiumXlsx -Destination $tempCopy -Force

  $wb = $excel.Workbooks.Open($tempCopy, 0, $true)

  $sheetIndex = 0
  if ([int]::TryParse($Sheet, [ref]$sheetIndex)) {
    if ($sheetIndex -lt 1 -or $sheetIndex -gt $wb.Sheets.Count) {
      throw "Sheet index $sheetIndex out of range (1..$($wb.Sheets.Count))."
    }
    $ws = $wb.Sheets.Item($sheetIndex)
  } else {
    $ws = $wb.Sheets.Item($Sheet)
  }
  $resolvedSheetName = $ws.Name

  $usedRange = $ws.UsedRange
  $rowCount = [int]$usedRange.Rows.Count
  $colCount = [int]$usedRange.Columns.Count
  $values = $usedRange.Value2  # 2D object array, 1-indexed [row, col]

  # Detect header row (scan up to row 30 for a quantity-like cell + something
  # with "Designator"). Accept exact "Quantity" or short "Qty" so that both
  # the engineering BOM xlsx (Quantity) and the production BOM-folder xlsx
  # (Qty) parse from the same script.
  $headerRow = -1
  $colMap = @{}
  $headerScanLimit = [Math]::Min($rowCount, 30)
  for ($r = 1; $r -le $headerScanLimit; $r++) {
    $rowVals = @()
    for ($c = 1; $c -le $colCount; $c++) { $rowVals += "" + $values[$r, $c] }
    $hasQty = $rowVals | Where-Object { $_.Trim() -imatch '^(quantity|qty)$' }
    $hasDes = $rowVals | Where-Object { $_ -match '(?i)designator' }
    if ($hasQty -and $hasDes) {
      $headerRow = $r
      $colMap.Clear()
      for ($c = 1; $c -le $colCount; $c++) {
        $h = ("" + $values[$r, $c]).Trim()
        if ($h) { $colMap[$h] = $c }
      }
      break
    }
  }
  if ($headerRow -lt 0) {
    throw "Could not find a header row containing 'Quantity' and a 'Designator' column on sheet '$resolvedSheetName'."
  }

  function Find-AltiumCol {
    param(
      [string[]]$Patterns,
      [string]$ExcludeColumnNamePattern = ''
    )
    foreach ($pat in $Patterns) {
      foreach ($name in @($colMap.Keys | Sort-Object)) {
        if ($ExcludeColumnNamePattern -and ($name -imatch $ExcludeColumnNamePattern)) { continue }
        if ($name -imatch $pat) { return $colMap[$name] }
      }
    }
    return -1
  }

  $cQty  = Find-AltiumCol @('^Quantity$','^Qty$','quantity','^qty$')
  # Altium engineering BOM uses *Part_Number; some exports / sheets use "ITEM NUMBER"
  # (Arena-style) or "Internal Part Number".
  # Prefer WDTS / ITEM NUMBER columns; fall back to 01-PJM Part_Number when that is
  # the only part-number column on the sheet (data rows still use it; header rows are skipped).
  $cPN   = Find-AltiumCol -Patterns @(
    '^WDTS\s*Part[_ ]?Number$',
    '^PN$',
    '^ITEM\s*NUMBER$',
    'Internal\s*Part\s*Number',
    'Item\s*Number',
    'Part[_ ]?Number'
  ) -ExcludeColumnNamePattern '(?i)^01-PJM\s+Part'
  if ($cPN -lt 1) {
    $cPN = Find-AltiumCol -Patterns @('Part[_ ]?Number', '^01-PJM\s+Part[_ ]?Number$')
  }
  $cDesc = Find-AltiumCol @('Description General','^Description')
  $cDes  = Find-AltiumCol @('^Designator$','Designators?','Reference\s*Designator')
  $cCmt  = Find-AltiumCol @('^Comment$')

  if ($cQty -lt 1 -or $cPN -lt 1 -or $cDes -lt 1) {
    throw ("Could not find Quantity / part-number / designator columns on Altium sheet '{0}'. Headers seen: {1}" -f `
      $resolvedSheetName, ((@($colMap.Keys)) -join ', '))
  }

  for ($r = $headerRow + 1; $r -le $rowCount; $r++) {
    $pn   = ("" + $values[$r, $cPN]).Trim()
    $qtyS = "" + $values[$r, $cQty]
    $qty = 0
    [int]::TryParse($qtyS, [ref]$qty) | Out-Null
    $desc = if ($cDesc -gt 0) { ("" + $values[$r, $cDesc]).Trim() } else { '' }
    $des  = Split-Designators ("" + $values[$r, $cDes])
    $cmt  = if ($cCmt -gt 0) { ("" + $values[$r, $cCmt]).Trim() } else { '' }

    if (-not $pn) {
      if ($des.Count -gt 0 -or $cmt) {
        $noPnRows += [pscustomobject]@{
          Quantity = $qty; Description = $desc; Comment = $cmt; Designators = $des
        }
      }
      continue
    }
    if (Test-IsAltiumBomNonDataPartNumberRow -PartNumber $pn -AssemblyPartNumber $AssemblyPartNumber `
        -Quantity $qty -Designators $des) {
      continue
    }
    $altium[$pn] = [pscustomobject]@{
      Quantity = $qty; Description = $desc; Comment = $cmt; Designators = $des
    }
  }

  if (-not [string]::IsNullOrWhiteSpace($PjmSpecificationSheet)) {
    $pj = Get-PjmSheetPartDescriptionRows -Workbook $wb -SheetSpec $PjmSpecificationSheet.Trim()
    $pjmDupPnGroups = @($pj.Rows | Where-Object { $_.PartNumber } | Group-Object { $_.PartNumber.Trim() } | Where-Object { $_.Count -gt 1 })
    foreach ($pr in $pj.Rows) {
      $pjn = $pr.PartNumber
      if (-not $pjn) { continue }
      if (-not $arena.ContainsKey($pjn)) { continue }
      $chk = Test-BomDescriptionsCompatible -ArenaDescription $arena[$pjn].Description -AltiumDescription $pr.Description -PartNumber $pjn
      if (-not $chk.Ok) {
        $pjmMismatches.Add([pscustomobject]@{
            PartNumber = $pjn
            Row        = $pr.Row
            Message    = $chk.Message
            Arena      = $arena[$pjn].Description
            Altium     = $pr.Description
            PjmSheet   = $pj.SheetName
          })
      }
    }
  }
} finally {
  if ($wb) {
    try { $wb.Close($false) } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
  }
  if ($excel) {
    try { $excel.Quit() } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  if ($tempCopy -and (Test-Path -LiteralPath $tempCopy)) {
    Remove-Item -LiteralPath $tempCopy -Force -ErrorAction SilentlyContinue
  }
}

# --------------------------- Diff ---------------------------
$onlyArena  = @($arena.Keys  | Where-Object { -not $altium.ContainsKey($_) } | Sort-Object)
$onlyAltium = @($altium.Keys | Where-Object { -not $arena.ContainsKey($_)  } | Sort-Object)
$common     = @($arena.Keys  | Where-Object { $altium.ContainsKey($_)      } | Sort-Object)

Write-Host ''
Write-Host "Arena CSV   : $ArenaCsv"
Write-Host "Altium xlsx : $AltiumXlsx"
Write-Host "Sheet       : $resolvedSheetName  (Sheet=$Sheet)"
if (-not [string]::IsNullOrWhiteSpace($AssemblyPartNumber)) {
  Write-Host ("Assembly    : {0} (excluded from Altium-only / common diff)" -f $AssemblyPartNumber.Trim())
}
Write-Host ("Arena parts : {0}    Altium parts: {1}    Common: {2}" -f $arena.Count, $altium.Count, $common.Count)

Write-Section "Parts only in Arena ($($onlyArena.Count))" $(if ($onlyArena.Count -eq 0) { 'DarkGreen' } else { 'DarkRed' })
if ($onlyArena.Count -eq 0) {
  Write-Host '  (none)' -ForegroundColor DarkGreen
} else {
  foreach ($pn in $onlyArena) {
    $a = $arena[$pn]
    Write-Host ("  {0,-12}  qty={1,-4}  des=[{2}]  -- {3}" -f $pn, $a.Quantity, ($a.Designators -join ', '), $a.Description) -ForegroundColor DarkRed
  }
}

Write-Section "Parts only in Altium ($($onlyAltium.Count))" $(if ($onlyAltium.Count -eq 0) { 'DarkGreen' } else { 'DarkRed' })
if ($onlyAltium.Count -eq 0) {
  Write-Host '  (none)' -ForegroundColor DarkGreen
} else {
  foreach ($pn in $onlyAltium) {
    $a = $altium[$pn]
    Write-Host ("  {0,-12}  qty={1,-4}  des=[{2}]  -- {3}" -f $pn, $a.Quantity, ($a.Designators -join ', '), $a.Description) -ForegroundColor DarkRed
  }
}

if ($noPnRows.Count -gt 0) {
  Write-Section "Altium rows without a part number ($($noPnRows.Count))" 'DarkYellow'
  foreach ($r in $noPnRows) {
    $tag = if ($r.Comment) { $r.Comment } else { '(blank)' }
    Write-Host ("  qty={0,-4}  des=[{1}]  -- {2}" -f $r.Quantity, ($r.Designators -join ', '), $tag)
  }
}

Write-Section "Common parts: quantity / designator differences" 'Cyan'
$qtyDesigDiffsFound = $false
foreach ($pn in $common) {
  $a = $arena[$pn]; $b = $altium[$pn]
  $diffs = @()
  if ($a.Quantity -ne $b.Quantity) {
    $diffs += "qty arena=$($a.Quantity) altium=$($b.Quantity)"
  }
  $aSet = New-Object System.Collections.Generic.HashSet[string] (,[string[]]$a.Designators)
  $bSet = New-Object System.Collections.Generic.HashSet[string] (,[string[]]$b.Designators)
  $onlyA = @($a.Designators | Where-Object { -not $bSet.Contains($_) })
  $onlyB = @($b.Designators | Where-Object { -not $aSet.Contains($_) })
  if ($onlyA.Count -gt 0) { $diffs += "designators only in Arena: [$($onlyA -join ', ')]" }
  if ($onlyB.Count -gt 0) { $diffs += "designators only in Altium: [$($onlyB -join ', ')]" }
  if ($diffs.Count -gt 0) {
    $qtyDesigDiffsFound = $true
    Write-Host "  $pn :" -ForegroundColor DarkRed
    foreach ($d in $diffs) { Write-Host "      $d" -ForegroundColor DarkRed }
  }
}
if (-not $qtyDesigDiffsFound) {
  Write-Host '  (no quantity or designator mismatches across common part numbers)' -ForegroundColor DarkGreen
}

Write-Section "Common parts: Arena item name vs Altium description (spec-aware)" 'Cyan'
$descSpecDiffsFound = $false
foreach ($pn in $common) {
  $chk = Test-BomDescriptionsCompatible -ArenaDescription $arena[$pn].Description -AltiumDescription $altium[$pn].Description -PartNumber $pn
  if (-not $chk.Ok) {
    $descSpecDiffsFound = $true
    Write-Host "  $pn :" -ForegroundColor DarkRed
    Write-Host ("      {0}" -f $chk.Message) -ForegroundColor DarkRed
    Write-Host ("      arena : {0}" -f $arena[$pn].Description) -ForegroundColor DarkRed
    Write-Host ("      altium: {0}" -f $altium[$pn].Description) -ForegroundColor DarkRed
  }
}
if (-not $descSpecDiffsFound) {
  Write-Host '  (no description / spec mismatches on the primary BOM sheet for common parts)' -ForegroundColor DarkGreen
}

if (-not [string]::IsNullOrWhiteSpace($PjmSpecificationSheet)) {
  Write-Section 'Duplicate 01-PJM Part_Number on PJM sheet' 'DarkYellow'
  if ($pjmDupPnGroups.Count -eq 0) {
    Write-Host '  (no duplicate part numbers on the PJM export tab)' -ForegroundColor DarkGreen
  }
  else {
    foreach ($g in $pjmDupPnGroups) {
      $rowNums = @($g.Group | ForEach-Object { $_.Row })
      Write-Host ("  '{0}'  Excel rows: {1}" -f $g.Name, ($rowNums -join ', ')) -ForegroundColor DarkRed
    }
  }

  $pjmTitle = "Sheet '$($PjmSpecificationSheet.Trim())' (01-PJM Part_Number vs Arena item name)"
  Write-Section $pjmTitle 'Cyan'
  if ($pjmMismatches.Count -eq 0) {
    Write-Host '  (all PJM export rows with a matching Arena line passed description checks)' -ForegroundColor DarkGreen
  }
  else {
    foreach ($m in $pjmMismatches) {
      Write-Host ("  {0} (excel row {1}) on '{2}':" -f $m.PartNumber, $m.Row, $m.PjmSheet)
      Write-Host ("      {0}" -f $m.Message)
      Write-Host ("      arena : {0}" -f $m.Arena)
      Write-Host ("      altium: {0}" -f $m.Altium)
    }
  }
}

$exitCode = 0
if ($descSpecDiffsFound) { $exitCode = 1 }
if ($pjmMismatches.Count -gt 0) { $exitCode = 1 }

Write-Host ''
if ($exitCode -eq 0) {
  Write-Host 'Done.' -ForegroundColor DarkGreen
}
else {
  Write-Host 'Done with description / PJM mismatches (see above).' -ForegroundColor DarkRed
}
exit $exitCode
