<#
.SYNOPSIS
  Export an Arena assembly BOM to import-style CSV, retarget 1206 tuning caps, and merge rows.

.DESCRIPTION
  1. Fetches the Arena BOM for the chosen revision (working or most recent released).
  2. Writes an import-format CSV with columns: LEVEL, Quantity, ITEM NUMBER, Reference Designator.
     Line 2 is the assembly row (LEVEL 0, empty Quantity and Reference Designator, ITEM NUMBER = PCBA).
  3. Reads a two-column designator / value sheet (Excel or CSV) using the same rules as
     Replace 1206 HV NP0 caps on PCBA BOM (pF -> Arena_1206_HighVolt_NP0_Caps.csv;
     TBD -> 098-13-987; N/F, NF, DNF, DNL -> 098-13-997; 0R -> 098-12-005).
  4. Updates ITEM NUMBER for listed designators and adds any missing designators.
  5. Merges rows that share the same ITEM NUMBER (combined Reference Designator, summed Quantity).

  Default output: Desktop\<PartNumber>_Bom1206TuningCaps_yyyyMMdd_HHmmss.csv

.PARAMETER PartNumber
  PCBA assembly part number.

.PARAMETER TuningSheetPath
  Excel (.xlsx/.xls) or CSV with designator and replacement value columns.

.PARAMETER RevisionMode
  working = BOM on working revision; effective_latest = newest released revision.

.PARAMETER OutputPath
  Optional full path for the output CSV.

.PARAMETER CapListPath
  CSV with PartNumber and ItemName (default: Arena_1206_HighVolt_NP0_Caps.csv beside this script).
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [Parameter(Mandatory = $true)] [string]$TuningSheetPath,
  [ValidateSet('working', 'effective_latest')] [string]$RevisionMode = 'working',
  [string]$OutputPath = '',
  [string]$CapListPath = '',
  [string]$Email = $env:ARENA_EMAIL,
  [string]$WorkspaceId = $env:ARENA_WORKSPACE_ID,
  [string]$BaseUrl = 'https://api.arenasolutions.com/v1'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
if ([string]::IsNullOrWhiteSpace($CapListPath)) {
  $CapListPath = Join-Path $scriptsDir 'Arena_1206_HighVolt_NP0_Caps.csv'
}

if ([string]::IsNullOrWhiteSpace($Email)) { throw 'Set ARENA_EMAIL.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_PASSWORD)) { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($WorkspaceId)) { throw 'Set ARENA_WORKSPACE_ID.' }
if (-not (Test-Path -LiteralPath $CapListPath)) { throw "Cap list not found: $CapListPath" }
if (-not (Test-Path -LiteralPath $TuningSheetPath)) { throw "Tuning sheet not found: $TuningSheetPath" }

$Script:SpecialPartTbd = '098-13-987'
$Script:SpecialPartNotFitted = '098-13-997'
$Script:SpecialPartZeroOhm = '098-12-005'
$Script:BomLevel = '1'

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
  try {
    $resp = Invoke-WebRequest @params
  } catch {
    $msg = $_.Exception.Message
    $bodyText = ''
    try {
      $errResp = $_.Exception.Response
      if ($errResp) {
        $stream = $errResp.GetResponseStream()
        if ($stream) {
          if ($stream.CanSeek) { $stream.Position = 0 }
          $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
          $bodyText = $reader.ReadToEnd()
          $reader.Dispose()
        }
      }
    } catch {}
    if ($bodyText) { $msg += " Response: $bodyText" }
    throw "${Method} ${Uri} failed: $msg"
  }
  if ($Method -eq 'Delete') { return $null }
  $bytes = $resp.RawContentStream.ToArray()
  $text = [System.Text.Encoding]::UTF8.GetString($bytes)
  if ([string]::IsNullOrWhiteSpace($text)) { return $null }
  return $text | ConvertFrom-Json
}

function Split-RefDesignators([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return @() }
  return @(($s -split '[,;]') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Parse-PfValue([string]$text) {
  if ([string]::IsNullOrWhiteSpace($text)) { return $null }
  $n = ("$text").Trim().ToUpperInvariant()
  $n = $n -replace '\s+', ''
  if ($n -match '^([\d.]+)UF$') { return [double]$Matches[1] * 1000000.0 }
  if ($n -match '^([\d.]+)NF$') { return [double]$Matches[1] * 1000.0 }
  if ($n -match '^([\d.]+)PF$') { return [double]$Matches[1] }
  if ($n -match '^([\d.]+)$') { return [double]$Matches[1] }
  return $null
}

function Resolve-ReplacementValue([string]$text) {
  if ([string]::IsNullOrWhiteSpace($text)) { return $null }
  $raw = ("$text").Trim()
  $compact = ($raw.ToUpperInvariant() -replace '\s+', '')
  $slashNorm = ($raw.ToUpperInvariant() -replace '\s*/\s*', '/').Trim()

  if ($compact -eq 'TBD') {
    return [pscustomobject]@{ Kind = 'part'; PartNumber = $Script:SpecialPartTbd; ValueLabel = $raw; Pf = $null }
  }
  $nfTokens = @('N/F', 'NF', 'DNF', 'DNL')
  if ($nfTokens -contains $slashNorm -or $nfTokens -contains $compact) {
    return [pscustomobject]@{ Kind = 'part'; PartNumber = $Script:SpecialPartNotFitted; ValueLabel = $raw; Pf = $null }
  }
  if ($compact -eq '0R') {
    return [pscustomobject]@{ Kind = 'part'; PartNumber = $Script:SpecialPartZeroOhm; ValueLabel = $raw; Pf = $null }
  }
  $pf = Parse-PfValue $raw
  if ($null -ne $pf) {
    return [pscustomobject]@{ Kind = 'pf'; PartNumber = $null; ValueLabel = $raw; Pf = $pf }
  }
  return $null
}

function Get-CapPfFromItemName([string]$name) {
  $n = ("$name").ToUpperInvariant()
  if ($n -match '([\d.]+)\s*UF\b') { return [double]$Matches[1] * 1000000.0 }
  if ($n -match '([\d.]+)\s*NF\b') { return [double]$Matches[1] * 1000.0 }
  if ($n -match '([\d.]+)\s*PF\b') { return [double]$Matches[1] }
  return $null
}

function Test-PfEqual([double]$a, [double]$b) {
  return [Math]::Abs($a - $b) -lt 0.001
}

function Find-ColumnIndexFromNames {
  param(
    [Parameter(Mandatory = $true)] [string[]]$ColumnNames,
    [Parameter(Mandatory = $true)] [string[]]$Candidates
  )
  $names = @($ColumnNames | ForEach-Object { "$_" })
  foreach ($cand in $Candidates) {
    $want = $cand.Trim().ToLowerInvariant()
    foreach ($name in $names) {
      if ("$name".Trim().ToLowerInvariant() -eq $want) { return $name }
    }
  }
  return $null
}

function Read-TuningReplacementRowsFromCsv {
  param([Parameter(Mandatory = $true)] [string]$Path)
  $rows = @(Import-Csv -LiteralPath $Path)
  if (-not $rows -or $rows.Count -eq 0) { throw 'Tuning CSV has no data rows.' }
  $cols = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
  $desCol = Find-ColumnIndexFromNames -ColumnNames $cols -Candidates @(
    'designator', 'designators', 'refdes', 'ref des', 'reference designator', 'reference designators'
  )
  $valCol = Find-ColumnIndexFromNames -ColumnNames $cols -Candidates @(
    'pf', 'pf value', 'picofarad', 'picofarads', 'capacitance', 'replacement', 'replacement pf', 'value'
  )
  if (-not $desCol -or -not $valCol) {
    if ($cols.Count -ge 2) {
      $desCol = $cols[0]
      $valCol = $cols[1]
    } else {
      throw 'Tuning CSV needs designator and value columns.'
    }
  }

  $out = New-Object System.Collections.Generic.List[object]
  foreach ($row in $rows) {
    $des = ("$($row.$desCol)").Trim()
    $valRaw = $row.$valCol
    if ([string]::IsNullOrWhiteSpace($des)) { continue }
    if ($des -match '(?i)^(designator|ref)') { continue }
    $resolved = Resolve-ReplacementValue ("$valRaw")
    if ($null -eq $resolved) {
      throw "Could not parse replacement value '$valRaw' for designator '$des'."
    }
    [void]$out.Add([pscustomobject]@{
        Designator = $des
        ValueLabel = $resolved.ValueLabel
        Kind       = $resolved.Kind
        Pf         = $resolved.Pf
        PartNumber = $resolved.PartNumber
      })
  }
  if ($out.Count -eq 0) { throw 'No designator / value rows found in tuning CSV.' }
  return ,$out.ToArray()
}

function Find-ExcelColumnIndex {
  param(
    [Parameter(Mandatory = $true)] $Worksheet,
    [Parameter(Mandatory = $true)] [int]$HeaderRow,
    [Parameter(Mandatory = $true)] [string[]]$Candidates
  )
  $used = $Worksheet.UsedRange
  if ($null -eq $used) { return $null }
  $endCol = $used.Column + $used.Columns.Count - 1
  for ([int]$c = 1; $c -le $endCol; $c++) {
    $raw = $Worksheet.Cells.Item($HeaderRow, $c).Value2
    if ($null -eq $raw) { continue }
    $cell = ("$raw").Trim().ToLowerInvariant()
    foreach ($want in $Candidates) {
      if ($cell -eq $want.ToLowerInvariant()) { return $c }
    }
  }
  return $null
}

function Read-TuningReplacementRowsFromExcel {
  param([Parameter(Mandatory = $true)] [string]$Path)
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $null
  try {
    $wb = $excel.Workbooks.Open((Resolve-Path -LiteralPath $Path).Path, 0, $true)
    $sheet = $wb.Worksheets.Item(1)
    $used = $sheet.UsedRange
    if ($null -eq $used) { throw 'Tuning Excel sheet is empty.' }
    $startRow = $used.Row
    $endRow = $used.Row + $used.Rows.Count - 1

    $desCol = Find-ExcelColumnIndex -Worksheet $sheet -HeaderRow $startRow -Candidates @(
      'designator', 'designators', 'refdes', 'ref des', 'reference designator', 'reference designators'
    )
    $valCol = Find-ExcelColumnIndex -Worksheet $sheet -HeaderRow $startRow -Candidates @(
      'pf', 'pf value', 'picofarad', 'picofarads', 'capacitance', 'replacement', 'replacement pf', 'value'
    )
    $dataStart = $startRow
    if ($desCol -and $valCol) {
      $dataStart = $startRow + 1
    } else {
      $desCol = $used.Column
      $valCol = $used.Column + 1
    }

    $out = New-Object System.Collections.Generic.List[object]
    for ([int]$r = $dataStart; $r -le $endRow; $r++) {
      $des = ("$($sheet.Cells.Item($r, $desCol).Value2)").Trim()
      $valRaw = $sheet.Cells.Item($r, $valCol).Value2
      if ([string]::IsNullOrWhiteSpace($des)) { continue }
      if ($des -match '(?i)^(designator|ref)') { continue }
      $resolved = Resolve-ReplacementValue ("$valRaw")
      if ($null -eq $resolved) {
        throw "Row $r ($des): could not parse replacement value '$valRaw'."
      }
      [void]$out.Add([pscustomobject]@{
          Designator = $des
          ValueLabel = $resolved.ValueLabel
          Kind       = $resolved.Kind
          Pf         = $resolved.Pf
          PartNumber = $resolved.PartNumber
        })
    }
    if ($out.Count -eq 0) { throw 'No designator / value rows found in tuning Excel.' }
    return ,$out.ToArray()
  }
  finally {
    if ($wb) {
      try { $wb.Close($false) } catch {}
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
    if ($excel) {
      try { $excel.Quit() } catch {}
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

function Read-TuningReplacementRows {
  param([Parameter(Mandatory = $true)] [string]$Path)
  $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
  if ($ext -eq '.csv') {
    return Read-TuningReplacementRowsFromCsv -Path $Path
  }
  if ($ext -in @('.xlsx', '.xls', '.xlsm')) {
    return Read-TuningReplacementRowsFromExcel -Path $Path
  }
  throw "Unsupported tuning sheet type '$ext' (use .csv, .xlsx, or .xls)."
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

function Build-TuningPlans {
  param(
    [Parameter(Mandatory = $true)] [object[]]$TuningRows,
    [Parameter(Mandatory = $true)] $PfToPart
  )
  $seenDes = @{}
  $plans = New-Object System.Collections.Generic.List[object]
  foreach ($row in $TuningRows) {
    $dk = $row.Designator.ToUpperInvariant()
    if ($seenDes.ContainsKey($dk)) {
      throw "Duplicate designator '$($row.Designator)' in tuning sheet."
    }
    $seenDes[$dk] = $true

    $targetPn = $null
    if ($row.Kind -eq 'part') {
      $targetPn = $row.PartNumber
    } else {
      $match = $null
      foreach ($entry in $PfToPart.Values) {
        if (Test-PfEqual $row.Pf $entry.Pf) { $match = $entry; break }
      }
      if (-not $match) {
        throw "No cap list part matches $($row.Pf) pF for designator $($row.Designator) (sheet value: $($row.ValueLabel))."
      }
      $targetPn = $match.PartNumber
    }
    [void]$plans.Add([pscustomobject]@{
        Designator    = $row.Designator
        ValueLabel    = $row.ValueLabel
        NewPartNumber = $targetPn
      })
  }
  return ,$plans.ToArray()
}

function ConvertTo-BomImportCsvLine {
  param(
    [string]$Level,
    [string]$Quantity = '',
    [string]$ItemNumber = '',
    [string]$RefDes = ''
  )
  $fields = @(
    $Level,
    $Quantity,
    $ItemNumber,
    $RefDes
  )
  $escaped = foreach ($f in $fields) {
    $s = if ($null -eq $f) { '' } else { "$f" }
    if ($s -match '[",\r\n]') { '"' + ($s -replace '"', '""') + '"' } else { $s }
  }
  return ($escaped -join ',')
}

function Merge-BomEntriesByPartNumber {
  param([Parameter(Mandatory = $true)] [object[]]$Entries)
  $groups = @{}
  foreach ($entry in $Entries) {
    $pn = $entry.PartNumber.Trim()
    if (-not $groups.ContainsKey($pn)) {
      $groups[$pn] = [pscustomobject]@{
        Level        = $entry.Level
        PartNumber   = $pn
        Designators  = New-Object System.Collections.Generic.List[string]
      }
    }
    [void]$groups[$pn].Designators.Add($entry.Designator)
  }

  $merged = New-Object System.Collections.Generic.List[object]
  foreach ($pn in ($groups.Keys | Sort-Object)) {
    $g = $groups[$pn]
    $desArray = $g.Designators.ToArray()
    [void]$merged.Add([pscustomobject]@{
        Level      = $g.Level
        Quantity   = $desArray.Count
        PartNumber = $g.PartNumber
        RefDes     = ($desArray -join ',')
      })
  }
  return ,$merged.ToArray()
}

# ---------- cap list ----------

Write-Host '=== Cap replacement list ==='
$capRows = Import-Csv -LiteralPath $CapListPath
$pfToPart = @{}
foreach ($row in $capRows) {
  $pf = Get-CapPfFromItemName $row.ItemName
  if ($null -eq $pf) {
    throw "Could not parse pF from cap list item $($row.PartNumber): $($row.ItemName)"
  }
  $key = [string]$pf
  if ($pfToPart.ContainsKey($key)) {
    throw "Duplicate pF $pf in cap list: $($pfToPart[$key].PartNumber) and $($row.PartNumber)."
  }
  $pfToPart[$key] = [pscustomobject]@{
    PartNumber = $row.PartNumber.Trim()
    ItemName   = $row.ItemName
    Pf         = $pf
  }
}
Write-Host ("  Loaded {0} cap part(s) from {1}" -f $capRows.Count, $CapListPath)

Write-Host ''
Write-Host '=== Tuning sheet ==='
$tuningRows = Read-TuningReplacementRows -Path $TuningSheetPath
$plans = Build-TuningPlans -TuningRows $tuningRows -PfToPart $pfToPart
$planByDes = @{}
foreach ($plan in $plans) {
  $planByDes[$plan.Designator.ToUpperInvariant()] = $plan
}
Write-Host ("  Rows: {0}  file: {1}" -f $plans.Count, $TuningSheetPath)

# ---------- Arena BOM ----------

Write-Host ''
Write-Host '=== Arena login ==='
$apiRoot = $BaseUrl.TrimEnd('/')
$loginBody = @{ email = $Email; password = $env:ARENA_PASSWORD; workspaceId = [long]$WorkspaceId } | ConvertTo-Json
$login = Invoke-ArenaJson -Uri "$apiRoot/login" -Method Post -Body $loginBody
$sid = if ($login.arenaSessionId) { $login.arenaSessionId } else { $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sid)) { throw 'Login response did not include arenaSessionId.' }
$headers = @{ arena_session_id = $sid; 'Content-Type' = 'application/json' }

$pn = $PartNumber.Trim()
$pnEnc = [Uri]::EscapeDataString($pn)
$lookup = Invoke-ArenaJson -Uri "$apiRoot/items?number=$pnEnc&limit=1" -Headers $headers
if (-not $lookup.results -or $lookup.results.Count -eq 0) {
  throw "No Arena item found for part number '$pn'."
}
$item = $lookup.results[0]
$revs = Invoke-ArenaJson -Uri "$apiRoot/items/$($item.guid)/revisions" -Headers $headers
if (-not $revs.results -or $revs.results.Count -eq 0) {
  throw "No revisions returned for Arena item '$pn'."
}
$assemblyRev = Select-ArenaAssemblyRevision -RevResults @($revs.results) -ModeKey $RevisionMode
$assemblyRevGuid = ("$($assemblyRev.guid)").Trim()
$revLabel = if ($RevisionMode -eq 'working') { 'working' } else { 'most recent released' }
Write-Host ("  Item: {0}  {1}" -f $item.number, $item.name)
Write-Host ("  BOM revision ({0}): {1}  guid={2}" -f $revLabel, $(if ($assemblyRev.number) { $assemblyRev.number } else { '(no number)' }), $assemblyRevGuid)

$bom = Invoke-ArenaJson -Uri "$apiRoot/items/$assemblyRevGuid/bom" -Headers $headers
$bomLines = if ($bom -and $bom.results) { @($bom.results) } else { @() }
Write-Host ("  BOM lines: {0}" -f $bomLines.Count)

# ---------- expand, retarget, merge ----------

$entries = New-Object System.Collections.Generic.List[object]
$desOnBom = @{}
foreach ($line in $bomLines) {
  if (-not $line.item -or -not $line.item.number) { continue }
  $partNumber = ("$($line.item.number)").Trim()
  $tokens = @(Split-RefDesignators ("$($line.refDes)"))
  if ($tokens.Count -eq 0) {
    [void]$entries.Add([pscustomobject]@{
        Level      = $Script:BomLevel
        Designator = ''
        PartNumber = $partNumber
      })
    continue
  }
  foreach ($des in $tokens) {
    $dk = $des.ToUpperInvariant()
    if ($desOnBom.ContainsKey($dk)) {
      throw "Designator '$des' appears on more than one Arena BOM line."
    }
    $desOnBom[$dk] = $true
    $pnOut = $partNumber
    if ($planByDes.ContainsKey($dk)) {
      $pnOut = $planByDes[$dk].NewPartNumber
    }
    [void]$entries.Add([pscustomobject]@{
        Level      = $Script:BomLevel
        Designator = $des
        PartNumber = $pnOut
      })
  }
}

foreach ($plan in $plans) {
  $dk = $plan.Designator.ToUpperInvariant()
  if ($desOnBom.ContainsKey($dk)) { continue }
  [void]$entries.Add([pscustomobject]@{
      Level      = $Script:BomLevel
      Designator = $plan.Designator
      PartNumber = $plan.NewPartNumber
    })
  $desOnBom[$dk] = $true
  Write-Host ("  Adding designator not on Arena BOM: {0} -> {1}" -f $plan.Designator, $plan.NewPartNumber)
}

$mergedRows = Merge-BomEntriesByPartNumber -Entries $entries.ToArray()

Write-Host ''
Write-Host '=== Planned tuning-cap changes ==='
Write-Host ("{0,-12} {1,-10} {2,-14} {3,-14} {4}" -f 'Designator', 'Value', 'Was PN', 'New PN', 'Status')
foreach ($plan in $plans) {
  $dk = $plan.Designator.ToUpperInvariant()
  $was = '(not on BOM)'
  $status = 'add'
  foreach ($line in $bomLines) {
    if (-not $line.item) { continue }
    foreach ($des in @(Split-RefDesignators ("$($line.refDes)"))) {
      if ($des.ToUpperInvariant() -eq $dk) {
        $was = ("$($line.item.number)").Trim()
        $status = if ($was -ieq $plan.NewPartNumber) { 'unchanged' } else { 'replace' }
        break
      }
    }
    if ($was -ne '(not on BOM)') { break }
  }
  Write-Host ("{0,-12} {1,-10} {2,-14} {3,-14} {4}" -f $plan.Designator, $plan.ValueLabel, $was, $plan.NewPartNumber, $status)
}

Write-Host ''
Write-Host '=== Output BOM CSV (merged by ITEM NUMBER) ==='
foreach ($row in $mergedRows) {
  $ref = if ($row.RefDes) { $row.RefDes } else { '' }
  Write-Host ("  {0}  qty={1}  {2}" -f $row.PartNumber, $row.Quantity, $ref)
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
  $desk = [Environment]::GetFolderPath('Desktop')
  if ([string]::IsNullOrWhiteSpace($desk)) {
    $desk = [Environment]::GetFolderPath('UserProfile')
  }
  $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
  $safePn = ($pn -replace '[^\w\-]+', '_')
  $OutputPath = Join-Path $desk ("{0}_Bom1206TuningCaps_{1}.csv" -f $safePn, $stamp)
}
else {
  $outDir = Split-Path -LiteralPath $OutputPath -Parent
  if (-not [string]::IsNullOrWhiteSpace($outDir) -and -not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  }
}

$csvLines = New-Object System.Collections.Generic.List[string]
[void]$csvLines.Add('LEVEL,Quantity,ITEM NUMBER,Reference Designator')
[void]$csvLines.Add((ConvertTo-BomImportCsvLine -Level '0' -ItemNumber $pn))
foreach ($row in $mergedRows) {
  [void]$csvLines.Add((ConvertTo-BomImportCsvLine -Level $row.Level -Quantity ("$($row.Quantity)") -ItemNumber $row.PartNumber -RefDes $row.RefDes))
}
$utf8Bom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllLines($OutputPath, $csvLines.ToArray(), $utf8Bom)

Write-Host ''
Write-Host ("Wrote {0} BOM line(s) to:" -f $mergedRows.Count) -ForegroundColor Green
Write-Host "  $OutputPath"
exit 0
