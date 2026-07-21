<#
.SYNOPSIS
  Replace PCBA BOM capacitors on the working revision from a designator / pF Excel sheet.

.DESCRIPTION
  Reads a two-column Excel sheet (designator, replacement value), maps each pF value to a part
  in scripts/Arena_1206_HighVolt_NP0_Caps.csv, and updates the working-revision BOM on the
  given PCBA assembly in Arena. Special replacement values: TBD -> 098-13-987;
  N/F, NF, DNF, or DNL -> 098-13-997; 0R -> 098-12-005. Designators not already on the
  working-revision BOM are
  added as new BOM lines.

  After updates, BOM lines with the same Arena part number are merged into one line (combined
  refDes, quantity = designator count). By default this is preview-only. Check "Apply changes
  to Arena BOM" in the launcher (or pass -ApplyChanges) to perform updates.

.PARAMETER PartNumber
  PCBA assembly part number (e.g. 272-15-001).

.PARAMETER ExcelPath
  Excel workbook (.xlsx/.xls) with designator and pF columns.

.PARAMETER CapListPath
  CSV with PartNumber and ItemName columns (default: Arena_1206_HighVolt_NP0_Caps.csv beside this script).

.PARAMETER ApplyChanges
  When set, apply BOM updates. Otherwise preview only.

.EXAMPLE
  .\Set-ArenaBomCapReplacementsFromSheet.ps1 -PartNumber 272-15-001 -ExcelPath .\caps.xlsx
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [Parameter(Mandatory = $true)] [string]$ExcelPath,
  [string]$CapListPath = '',
  [switch]$ApplyChanges,
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
if (-not (Test-Path -LiteralPath $ExcelPath)) { throw "Excel file not found: $ExcelPath" }

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
    if ($Body) { $msg += " Request: $Body" }
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

$Script:SpecialPartTbd = '098-13-987'
$Script:SpecialPartNotFitted = '098-13-997'
$Script:SpecialPartZeroOhm = '098-12-005'

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
    return [pscustomobject]@{
      Kind        = 'part'
      PartNumber  = $Script:SpecialPartTbd
      ValueLabel  = $raw
      Pf          = $null
    }
  }

  $nfTokens = @('N/F', 'NF', 'DNF', 'DNL')
  if ($nfTokens -contains $slashNorm -or $nfTokens -contains $compact) {
    return [pscustomobject]@{
      Kind        = 'part'
      PartNumber  = $Script:SpecialPartNotFitted
      ValueLabel  = $raw
      Pf          = $null
    }
  }

  if ($compact -eq '0R') {
    return [pscustomobject]@{
      Kind        = 'part'
      PartNumber  = $Script:SpecialPartZeroOhm
      ValueLabel  = $raw
      Pf          = $null
    }
  }

  $pf = Parse-PfValue $raw
  if ($null -ne $pf) {
    return [pscustomobject]@{
      Kind        = 'pf'
      PartNumber  = $null
      ValueLabel  = $raw
      Pf          = $pf
    }
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

function Read-ReplacementRowsFromExcel {
  param([Parameter(Mandatory = $true)] [string]$Path)
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $null
  try {
    $wb = $excel.Workbooks.Open((Resolve-Path -LiteralPath $Path).Path, 0, $true)
    $sheet = $wb.Worksheets.Item(1)
    $used = $sheet.UsedRange
    if ($null -eq $used) { throw 'Excel sheet is empty.' }
    $startRow = $used.Row
    $endRow = $used.Row + $used.Rows.Count - 1

    $desCol = Find-ExcelColumnIndex -Worksheet $sheet -HeaderRow $startRow -Candidates @(
      'designator', 'designators', 'refdes', 'ref des', 'reference designator', 'reference designators'
    )
    $pfCol = Find-ExcelColumnIndex -Worksheet $sheet -HeaderRow $startRow -Candidates @(
      'pf', 'pf value', 'picofarad', 'picofarads', 'capacitance', 'replacement', 'replacement pf', 'value'
    )
    $dataStart = $startRow
    if ($desCol -and $pfCol) {
      $dataStart = $startRow + 1
    } else {
      $desCol = $used.Column
      $pfCol = $used.Column + 1
    }

    $rows = New-Object System.Collections.Generic.List[object]
    for ([int]$r = $dataStart; $r -le $endRow; $r++) {
      $des = ("$($sheet.Cells.Item($r, $desCol).Value2)").Trim()
      $pfRaw = $sheet.Cells.Item($r, $pfCol).Value2
      if ([string]::IsNullOrWhiteSpace($des)) { continue }
      if ($des -match '(?i)^(designator|ref)') { continue }
      $resolved = Resolve-ReplacementValue ("$pfRaw")
      if ($null -eq $resolved) {
        throw "Row $r ($des): could not parse replacement value '$pfRaw' (expected pF, 0R, TBD, N/F, NF, DNF, or DNL)."
      }
      [void]$rows.Add([pscustomobject]@{
          Designator  = $des
          ValueRaw    = "$pfRaw"
          ValueLabel  = $resolved.ValueLabel
          Kind        = $resolved.Kind
          Pf          = $resolved.Pf
          PartNumber  = $resolved.PartNumber
        })
    }
    if ($rows.Count -eq 0) { throw 'No designator / pF rows found in Excel.' }
    return ,$rows.ToArray()
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

function Get-ArenaWorkingAssemblyGuid {
  param(
    [string]$ApiRoot,
    [hashtable]$Headers,
    [string]$Pn
  )
  $pnEnc = [Uri]::EscapeDataString($Pn.Trim())
  $lookup = Invoke-ArenaJson -Uri "$ApiRoot/items?number=$pnEnc&limit=1" -Headers $Headers
  if (-not $lookup.results -or $lookup.results.Count -eq 0) {
    throw "No Arena item found for part number '$Pn'."
  }
  $anyGuid = "$($lookup.results[0].guid)".Trim()
  $revs = Invoke-ArenaJson -Uri "$ApiRoot/items/$anyGuid/revisions" -Headers $Headers
  $working = $revs.results | Where-Object { $_.status -eq 0 } | Select-Object -First 1
  if (-not $working) { throw "No working revision (status 0) for '$Pn'." }
  return [pscustomobject]@{
    WorkingGuid = "$($working.guid)".Trim()
    RevNumber   = if ($null -ne $working.number) { "$($working.number)".Trim() } else { '' }
    Name        = $lookup.results[0].name
    Number      = $lookup.results[0].number
  }
}

function Get-ArenaItemGuidByNumber {
  param(
    [string]$ApiRoot,
    [hashtable]$Headers,
    [string]$Pn,
    [hashtable]$Cache
  )
  $key = $Pn.Trim()
  if ($Cache.ContainsKey($key)) { return $Cache[$key] }
  $pnEnc = [Uri]::EscapeDataString($key)
  $lookup = Invoke-ArenaJson -Uri "$ApiRoot/items?number=$pnEnc&limit=1" -Headers $Headers
  if (-not $lookup.results -or $lookup.results.Count -eq 0) {
    throw "No Arena item found for part number '$key'."
  }
  $guid = "$($lookup.results[0].guid)".Trim()
  $Cache[$key] = $guid
  return $guid
}

function Build-BomDesignatorIndex {
  param([Parameter(Mandatory = $true)] $BomResults)
  $index = @{}
  foreach ($line in @($BomResults)) {
    foreach ($des in @(Split-RefDesignators ("$($line.refDes)"))) {
      if (-not $des) { continue }
      $key = $des.ToUpperInvariant()
      if ($index.ContainsKey($key)) {
        throw "Designator '$des' appears on more than one BOM line."
      }
      $index[$key] = $line
    }
  }
  return $index
}

function New-BomLineBody {
  param(
    [string]$RefDes,
    [int]$Quantity,
    [string]$ItemGuid = '',
    [string]$Notes = '',
    [switch]$IncludeItem
  )
  $body = [ordered]@{
    refDes   = $RefDes
    quantity = $Quantity
  }
  if ($IncludeItem) {
    if ([string]::IsNullOrWhiteSpace($ItemGuid)) {
      throw 'ItemGuid is required when IncludeItem is set on a BOM line body.'
    }
    $body.item = @{ guid = $ItemGuid }
  }
  if ($null -ne $Notes -and "$Notes".Trim()) { $body.notes = "$Notes".Trim() }
  return ($body | ConvertTo-Json -Compress)
}

function Set-BomLine {
  param(
    [string]$ApiRoot,
    [hashtable]$Headers,
    [string]$AssemblyGuid,
    [string]$LineGuid,
    [string]$BodyJson
  )
  [void](Invoke-ArenaJson -Uri "$ApiRoot/items/$AssemblyGuid/bom/$LineGuid" -Method Put -Headers $Headers -Body $BodyJson)
}

function Add-BomLine {
  param(
    [string]$ApiRoot,
    [hashtable]$Headers,
    [string]$AssemblyGuid,
    [string]$BodyJson
  )
  [void](Invoke-ArenaJson -Uri "$ApiRoot/items/$AssemblyGuid/bom" -Method Post -Headers $Headers -Body $BodyJson)
}

function Remove-BomLine {
  param(
    [string]$ApiRoot,
    [hashtable]$Headers,
    [string]$AssemblyGuid,
    [string]$LineGuid
  )
  [void](Invoke-ArenaJson -Uri "$ApiRoot/items/$AssemblyGuid/bom/$LineGuid" -Method Delete -Headers $Headers)
}

function Apply-ReplacementsOnBomLine {
  param(
    [string]$ApiRoot,
    [hashtable]$Headers,
    [string]$AssemblyGuid,
    [Parameter(Mandatory = $true)] $BomLine,
    [Parameter(Mandatory = $true)] [object[]]$Replacements,
    [hashtable]$ItemGuidCache
  )
  $lineGuid = "$($BomLine.guid)".Trim()
  $oldItemGuid = "$($BomLine.item.guid)".Trim()
  $notes = if ($null -ne $BomLine.notes) { "$($BomLine.notes)" } else { '' }
  $allTokens = @(Split-RefDesignators ("$($BomLine.refDes)"))
  $replaceByDes = @{}
  foreach ($r in $Replacements) {
    $replaceByDes[$r.Designator.ToUpperInvariant()] = $r
  }

  $unchanged = @($allTokens | Where-Object { -not $replaceByDes.ContainsKey($_.ToUpperInvariant()) })
  $toReplace = @($allTokens | Where-Object { $replaceByDes.ContainsKey($_.ToUpperInvariant()) })

  if ($toReplace.Count -eq 0) { return 0 }

  $currentPn = ("$($BomLine.item.number)").Trim()
  $pending = New-Object System.Collections.Generic.List[object]
  foreach ($des in $toReplace) {
    $rep = $replaceByDes[$des.ToUpperInvariant()]
    if ($currentPn -ieq $rep.NewPartNumber) { continue }
    [void]$pending.Add([pscustomobject]@{ Designator = $des; Replacement = $rep })
  }
  if ($pending.Count -eq 0) { return 0 }

  $changed = 0
  if ($unchanged.Count -eq 0) {
    Remove-BomLine -ApiRoot $ApiRoot -Headers $Headers -AssemblyGuid $AssemblyGuid -LineGuid $lineGuid
  } else {
    $refDes = ($unchanged -join ',')
    $body = New-BomLineBody -RefDes $refDes -Quantity $unchanged.Count -Notes $notes
    Set-BomLine -ApiRoot $ApiRoot -Headers $Headers -AssemblyGuid $AssemblyGuid -LineGuid $lineGuid -BodyJson $body
  }

  $byPart = @{}
  foreach ($entry in $pending) {
    $pn = $entry.Replacement.NewPartNumber.Trim()
    if (-not $byPart.ContainsKey($pn)) {
      $byPart[$pn] = New-Object System.Collections.Generic.List[string]
    }
    [void]$byPart[$pn].Add($entry.Designator)
  }
  foreach ($pn in $byPart.Keys) {
    $designators = $byPart[$pn].ToArray()
    $newItemGuid = Get-ArenaItemGuidByNumber -ApiRoot $ApiRoot -Headers $Headers -Pn $pn -Cache $ItemGuidCache
    $refDes = ($designators -join ',')
    $body = New-BomLineBody -RefDes $refDes -Quantity $designators.Count -ItemGuid $newItemGuid -Notes $notes -IncludeItem
    Add-BomLine -ApiRoot $ApiRoot -Headers $Headers -AssemblyGuid $AssemblyGuid -BodyJson $body
    $changed++
  }
  return $changed
}

function Merge-ArenaBomLinesByPartNumber {
  param(
    [string]$ApiRoot,
    [hashtable]$Headers,
    [string]$AssemblyGuid
  )
  $bom = Invoke-ArenaJson -Uri "$ApiRoot/items/$AssemblyGuid/bom" -Headers $Headers
  if (-not $bom -or -not $bom.results) { return 0 }

  $groups = @{}
  foreach ($line in @($bom.results)) {
    if (-not $line.item -or -not $line.item.guid) { continue }
    $pn = ("$($line.item.number)").Trim()
    if (-not $pn) { continue }
    if (-not $groups.ContainsKey($pn)) {
      $groups[$pn] = New-Object System.Collections.Generic.List[object]
    }
    [void]$groups[$pn].Add($line)
  }

  $mergedGroups = 0
  foreach ($pn in @($groups.Keys)) {
    $lines = $groups[$pn].ToArray()
    if ($lines.Count -le 1) { continue }

    $seenDes = @{}
    $designators = New-Object System.Collections.Generic.List[string]
    $notes = ''
    $itemGuid = ''
    foreach ($line in $lines) {
      foreach ($des in @(Split-RefDesignators ("$($line.refDes)"))) {
        if (-not $des) { continue }
        $key = $des.ToUpperInvariant()
        if ($seenDes.ContainsKey($key)) { continue }
        $seenDes[$key] = $true
        [void]$designators.Add($des)
      }
      if (-not $notes -and $null -ne $line.notes -and "$($line.notes)".Trim()) {
        $notes = "$($line.notes)".Trim()
      }
      if (-not $itemGuid) { $itemGuid = "$($line.item.guid)".Trim() }
    }
    if ($designators.Count -eq 0 -or -not $itemGuid) { continue }

    foreach ($line in $lines) {
      Remove-BomLine -ApiRoot $ApiRoot -Headers $Headers -AssemblyGuid $AssemblyGuid -LineGuid ("$($line.guid)".Trim())
    }
    $desArray = $designators.ToArray()
    $refDes = ($desArray -join ',')
    $body = New-BomLineBody -RefDes $refDes -Quantity $desArray.Count -ItemGuid $itemGuid -Notes $notes -IncludeItem
    Add-BomLine -ApiRoot $ApiRoot -Headers $Headers -AssemblyGuid $AssemblyGuid -BodyJson $body
    $mergedGroups++
    Write-Host ("  Merged {0} BOM line(s) for {1}: {2}" -f $lines.Count, $pn, $refDes)
  }
  return $mergedGroups
}

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
    throw "Duplicate pF $pf in cap list: $($pfToPart[$key]) and $($row.PartNumber)."
  }
  $pfToPart[$key] = [pscustomobject]@{
    PartNumber = $row.PartNumber.Trim()
    ItemName   = $row.ItemName
    Pf         = $pf
  }
}
Write-Host ("  Loaded {0} cap part(s) from {1}" -f $capRows.Count, $CapListPath)

Write-Host ''
Write-Host '=== Excel replacements ==='
$excelRows = Read-ReplacementRowsFromExcel -Path $ExcelPath
Write-Host ("  Rows: {0}  file: {1}" -f $excelRows.Count, $ExcelPath)

$seenDes = @{}
foreach ($row in $excelRows) {
  $dk = $row.Designator.ToUpperInvariant()
  if ($seenDes.ContainsKey($dk)) {
    throw "Duplicate designator '$($row.Designator)' in the Excel sheet."
  }
  $seenDes[$dk] = $true
}

$plans = New-Object System.Collections.Generic.List[object]
foreach ($row in $excelRows) {
  if ($row.Kind -eq 'part') {
    [void]$plans.Add([pscustomobject]@{
        Designator    = $row.Designator
        ValueLabel    = $row.ValueLabel
        NewPartNumber = $row.PartNumber
        NewItemName   = ''
      })
    continue
  }

  $match = $null
  foreach ($entry in $pfToPart.Values) {
    if (Test-PfEqual $row.Pf $entry.Pf) {
      $match = $entry
      break
    }
  }
  if (-not $match) {
    throw ("No cap list part matches {0} pF for designator {1} (sheet value: {2})." -f $row.Pf, $row.Designator, $row.ValueRaw)
  }
  [void]$plans.Add([pscustomobject]@{
      Designator    = $row.Designator
      ValueLabel    = $row.ValueLabel
      NewPartNumber = $match.PartNumber
      NewItemName   = $match.ItemName
    })
}

Write-Host ''
Write-Host '=== Arena login ==='
$apiRoot = $BaseUrl.TrimEnd('/')
$loginBody = @{ email = $Email; password = $env:ARENA_PASSWORD; workspaceId = [long]$WorkspaceId } | ConvertTo-Json
$login = Invoke-ArenaJson -Uri "$apiRoot/login" -Method Post -Body $loginBody
$sid = if ($login.arenaSessionId) { $login.arenaSessionId } else { $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sid)) { throw 'Login response did not include arenaSessionId.' }
$headers = @{ arena_session_id = $sid; 'Content-Type' = 'application/json' }

$assembly = Get-ArenaWorkingAssemblyGuid -ApiRoot $apiRoot -Headers $headers -Pn $PartNumber.Trim()
Write-Host ("  PCBA: {0}  {1}" -f $assembly.Number, $assembly.Name)
Write-Host ("  Working revision: {0}  guid={1}" -f $(if ($assembly.RevNumber) { $assembly.RevNumber } else { '(no number)' }), $assembly.WorkingGuid)

$bom = Invoke-ArenaJson -Uri "$apiRoot/items/$($assembly.WorkingGuid)/bom" -Headers $headers
$bomLines = if ($bom -and $bom.results) { @($bom.results) } else { @() }
$desIndex = Build-BomDesignatorIndex -BomResults $bomLines
Write-Host ("  BOM lines on working revision: {0}" -f $bomLines.Count)

Write-Host ''
Write-Host '=== Planned replacements ==='
Write-Host ("{0,-12} {1,-10} {2,-14} {3,-14} {4}" -f 'Designator', 'Value', 'Current PN', 'New PN', 'Status')
$applyGroups = @{}
$addPlans = New-Object System.Collections.Generic.List[object]
foreach ($plan in $plans) {
  $desKey = $plan.Designator.ToUpperInvariant()
  if (-not $desIndex.ContainsKey($desKey)) {
    Write-Host ("{0,-12} {1,-10} {2,-14} {3,-14} {4}" -f $plan.Designator, $plan.ValueLabel, '(not on BOM)', $plan.NewPartNumber, 'add')
    [void]$addPlans.Add($plan)
    continue
  }
  $line = $desIndex[$desKey]
  $currentPn = ("$($line.item.number)").Trim()
  $status = if ($currentPn -ieq $plan.NewPartNumber) { 'unchanged' } else { 'replace' }
  Write-Host ("{0,-12} {1,-10} {2,-14} {3,-14} {4}" -f $plan.Designator, $plan.ValueLabel, $currentPn, $plan.NewPartNumber, $status)
  if ($status -eq 'replace') {
    $lineGuid = "$($line.guid)".Trim()
    if (-not $applyGroups.ContainsKey($lineGuid)) {
      $applyGroups[$lineGuid] = [pscustomobject]@{
        BomLine      = $line
        Replacements = New-Object System.Collections.Generic.List[object]
      }
    }
    [void]$applyGroups[$lineGuid].Replacements.Add($plan)
  }
}

$replaceCount = 0
foreach ($group in $applyGroups.Values) {
  $replaceCount += $group.Replacements.Count
}
$addCount = $addPlans.Count
$changeCount = $replaceCount + $addCount

if (-not $ApplyChanges) {
  if ($changeCount -eq 0) {
    Write-Host ''
    Write-Host 'All designators already use the matching cap parts. Nothing to do.' -ForegroundColor Green
    exit 0
  }
  Write-Host ''
  $previewParts = New-Object System.Collections.Generic.List[string]
  if ($replaceCount -gt 0) { [void]$previewParts.Add("$replaceCount to replace") }
  if ($addCount -gt 0) { [void]$previewParts.Add("$addCount to add") }
  Write-Host ("Preview only: {0}." -f ($previewParts -join ', ')) -ForegroundColor Yellow
  Write-Host 'Check "Apply changes to Arena BOM" and run again to update the working revision.'
  exit 0
}

Write-Host ''
Write-Host '=== Applying BOM updates ==='
$itemCache = @{}
$totalChanged = 0
if ($changeCount -eq 0) {
  Write-Host '  No designator value changes; consolidating BOM lines by part number only.'
}
foreach ($group in $applyGroups.Values) {
  $n = Apply-ReplacementsOnBomLine `
    -ApiRoot $apiRoot `
    -Headers $headers `
    -AssemblyGuid $assembly.WorkingGuid `
    -BomLine $group.BomLine `
    -Replacements $group.Replacements.ToArray() `
    -ItemGuidCache $itemCache
  $totalChanged += $n
}
$addByPart = @{}
if ($addPlans.Count -gt 0) {
foreach ($plan in $addPlans) {
  $pn = $plan.NewPartNumber.Trim()
  if (-not $addByPart.ContainsKey($pn)) {
    $addByPart[$pn] = New-Object System.Collections.Generic.List[string]
  }
  [void]$addByPart[$pn].Add($plan.Designator)
}
foreach ($pn in $addByPart.Keys) {
  $designators = $addByPart[$pn].ToArray()
  $newItemGuid = Get-ArenaItemGuidByNumber -ApiRoot $apiRoot -Headers $headers -Pn $pn -Cache $itemCache
  $refDes = ($designators -join ',')
  $body = New-BomLineBody -RefDes $refDes -Quantity $designators.Count -ItemGuid $newItemGuid -Notes '' -IncludeItem
  Add-BomLine -ApiRoot $apiRoot -Headers $headers -AssemblyGuid $assembly.WorkingGuid -BodyJson $body
  $totalChanged++
  Write-Host ("  Added {0} -> {1}" -f $refDes, $pn)
}
}

Write-Host ''
Write-Host '=== Consolidating BOM lines by part number ==='
$mergedGroups = Merge-ArenaBomLinesByPartNumber -ApiRoot $apiRoot -Headers $headers -AssemblyGuid $assembly.WorkingGuid
if ($mergedGroups -eq 0) {
  Write-Host '  No duplicate part-number lines to merge.'
}

Write-Host ''
Write-Host ("  Applied {0} BOM change(s) on working revision ({1} replace, {2} add); merged {3} part group(s)." -f `
  $totalChanged, $replaceCount, $addCount, $mergedGroups) -ForegroundColor Green
exit 0
