<#
.SYNOPSIS
  Preview then upload an Altium production BOM CSV to Arena (Import Engine ITEM_BOM).

.DESCRIPTION
  Given a part number:
    1. Resolves the Altium production BOM folder and picks the newest .csv there.
    2. Prints PCBA (PPP-10/15) and PCB (PPP-20/25) part numbers from the paired Excel
       export (newest .xlsx/.xls in the same folder), falling back to CSV lines 2-3.
    3. If the Arena assembly BOM has lines, runs the same comparison as
       Compare Arena vs Altium BOM v2; otherwise prints that the Arena BOM is empty.
    4. Without -ConfirmUpload, prints ---UPLOAD-CONFIRM-PROMPT--- and exits 12 so the
       Script Launcher shows a final Confirm Upload button.
    5. With -ConfirmUpload, runs an Arena Import Engine ITEM_BOM import for the CSV.

  Import mode (-ImportMode or ARENA_BOM_IMPORT_MODE):
    auto   = CREATE when Arena BOM is empty, REPLACE when it already has lines (default)
    CREATE = add BOM lines (typical for a new / empty working-rev BOM)
    REPLACE = replace the entire BOM view (only when you intend to wipe and reload)
    OVERWRITE / EDIT = use only if you have matching definitions in Arena

  If your workspace has an Import Definition (ITEM_BOM) for the chosen mode, that is
  used. Otherwise the script adds BOM lines via the Arena REST API (POST /items/.../bom),
  which does not require an import definition (typical for empty BOM / CREATE).

  Optional: -ImportDefinitionGuid or ARENA_BOM_IMPORT_DEFINITION_GUID to force a definition.
  Optional: -ForceDirectApi to skip Import Engine and always use REST line-by-line import.

  Credentials: ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID.

.PARAMETER PartNumber
  Part number for Altium folder resolve and Arena BOM lookup (e.g. 272-15-001).

.PARAMETER ConfirmUpload
  When set, perform the Arena BOM import after preview checks.

.PARAMETER ImportMode
  auto, CREATE, REPLACE, OVERWRITE, or EDIT. Default: auto or ARENA_BOM_IMPORT_MODE.

.PARAMETER ImportDefinitionGuid
  Arena Import Definition GUID (ITEM_BOM). Default: ARENA_BOM_IMPORT_DEFINITION_GUID.

.EXAMPLE
  .\Upload-ArenaBomFromAltiumCsv.ps1 -PartNumber 296-15-001 -ImportMode CREATE
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [switch]$ConfirmUpload,
  [ValidateSet('auto', 'CREATE', 'REPLACE', 'OVERWRITE', 'EDIT')]
  [string]$ImportMode = $(if ($env:ARENA_BOM_IMPORT_MODE) { $env:ARENA_BOM_IMPORT_MODE.Trim() } else { 'auto' }),
  [string]$ImportDefinitionGuid = $env:ARENA_BOM_IMPORT_DEFINITION_GUID,
  [switch]$ForceDirectApi,
  [string]$Email = $env:ARENA_EMAIL,
  [string]$WorkspaceId = $env:ARENA_WORKSPACE_ID,
  [string]$BaseUrl = 'https://api.arenasolutions.com/v1'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$openScript = Join-Path $scriptsDir 'Open-AltiumPartFolder.ps1'
$getBomScript = Join-Path $scriptsDir 'Get-ArenaBom.ps1'
$compareAutoScript = Join-Path $scriptsDir 'Compare-ArenaAltiumBomAuto.ps1'
foreach ($p in @($openScript, $getBomScript, $compareAutoScript)) {
  if (-not (Test-Path -LiteralPath $p)) { throw "Required helper script not found: $p" }
}

if ([string]::IsNullOrWhiteSpace($Email)) { throw 'Set ARENA_EMAIL.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_PASSWORD)) { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($WorkspaceId)) { throw 'Set ARENA_WORKSPACE_ID.' }

function Get-PartNumberFromBomText {
  param(
    [string]$Text,
    [string]$MiddleSegmentPattern
  )
  if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
  $m = [regex]::Match($Text, "(?i)(\d{3}-(?:$MiddleSegmentPattern)-\d{3})")
  if ($m.Success) { return $m.Groups[1].Value }
  return $null
}

function Get-PcbaPcbFromExcelBom {
  param([Parameter(Mandatory)][string]$ExcelPath)
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $wb = $null
  try {
    $wb = $excel.Workbooks.Open($ExcelPath, 0, $true)
    $sheet = $wb.Worksheets.Item(1)
    $pcba = $null
    $pcb = $null
    for ([int]$r = 1; $r -le 3; $r++) {
      $used = $sheet.UsedRange
      if ($null -eq $used) { break }
      $endCol = $used.Column + $used.Columns.Count - 1
      $lineText = ''
      for ([int]$c = 1; $c -le $endCol; $c++) {
        $raw = $sheet.Cells.Item($r, $c).Value2
        if ($null -ne $raw -and -not [string]::IsNullOrWhiteSpace(([string]$raw).Trim())) {
          $lineText += ' ' + ([string]$raw).Trim()
        }
      }
      if (-not $pcba) { $pcba = Get-PartNumberFromBomText -Text $lineText -MiddleSegmentPattern '10|15' }
      if (-not $pcb) { $pcb = Get-PartNumberFromBomText -Text $lineText -MiddleSegmentPattern '20|25' }
    }
    return [pscustomobject]@{ Pcba = $pcba; Pcb = $pcb }
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

function Get-ArenaApiErrorBodyText {
  param($ErrorRecord)
  if (-not $ErrorRecord) { return '' }
  $resp = $ErrorRecord.Exception.Response
  if (-not $resp) { return '' }
  try {
    $stream = $resp.GetResponseStream()
    if (-not $stream) { return '' }
    if ($stream.CanSeek) { $stream.Position = 0 }
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
    try {
      return $reader.ReadToEnd()
    } finally {
      $reader.Dispose()
    }
  } catch {
    return ''
  }
}

function Format-ArenaApiErrorForDisplay {
  param(
    [string]$Uri,
    [string]$Method,
    $ErrorRecord
  )
  $bodyText = Get-ArenaApiErrorBodyText -ErrorRecord $ErrorRecord
  $parts = New-Object System.Collections.Generic.List[string]
  [void]$parts.Add(("{0} {1} failed: {2}" -f $Method, $Uri, $ErrorRecord.Exception.Message))
  if ([string]::IsNullOrWhiteSpace($bodyText)) {
    return ($parts -join ' ')
  }
  [void]$parts.Add("Arena response body: $bodyText")
  try {
    $parsed = $bodyText | ConvertFrom-Json
    if ($parsed.errors) {
      foreach ($err in @($parsed.errors)) {
        $code = if ($err.code) { $err.code } else { '' }
        $msg = if ($err.message) { $err.message } else { "$err" }
        [void]$parts.Add(("  Arena error {0}: {1}" -f $code, $msg))
      }
    }
  } catch {
    # Body was not JSON; raw text already included above.
  }
  return ($parts -join [Environment]::NewLine)
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
  try {
    $resp = Invoke-WebRequest @params
  } catch {
    throw (Format-ArenaApiErrorForDisplay -Uri $Uri -Method $Method -ErrorRecord $_)
  }
  $bytes = $resp.RawContentStream.ToArray()
  $text = [System.Text.Encoding]::UTF8.GetString($bytes)
  if ([string]::IsNullOrWhiteSpace($text)) { return $null }
  return $text | ConvertFrom-Json
}

function Invoke-ArenaMultipartImportRun {
  param(
    [string]$ApiRoot,
    [hashtable]$Headers,
    [string]$ImportDefinitionGuid,
    [string]$CsvPath,
    [switch]$Commit
  )
  Add-Type -AssemblyName System.Net.Http -ErrorAction Stop
  $boundary = [System.Guid]::NewGuid().ToString('N')
  $enc = [System.Text.Encoding]::UTF8
  $LF = "`r`n"
  $bodyStream = New-Object System.IO.MemoryStream

  function Write-StringToStream([System.IO.MemoryStream]$Stream, [string]$Text) {
    $b = $enc.GetBytes($Text)
    $Stream.Write($b, 0, $b.Length)
  }
  function Add-Field([System.IO.MemoryStream]$Stream, [string]$Name, [string]$Value) {
    $part = "--$boundary$LF" +
            "Content-Disposition: form-data; name=`"$Name`"$LF$LF" +
            "$Value$LF"
    Write-StringToStream $Stream $part
  }
  function Add-File([System.IO.MemoryStream]$Stream, [string]$Name, [string]$FileName, [byte[]]$Bytes) {
    $header = "--$boundary$LF" +
              "Content-Disposition: form-data; name=`"$Name`"; filename=`"$FileName`"$LF" +
              "Content-Type: text/csv$LF$LF"
    Write-StringToStream $Stream $header
    $Stream.Write($Bytes, 0, $Bytes.Length)
    Write-StringToStream $Stream $LF
  }

  $fileName = [System.IO.Path]::GetFileName($CsvPath)
  $fileBytes = [System.IO.File]::ReadAllBytes($CsvPath)
  Add-Field $bodyStream 'submitFileType' 'CSV'
  Add-Field $bodyStream 'commit' ($(if ($Commit) { 'true' } else { 'false' }))
  Add-File $bodyStream 'submitContent' $fileName $fileBytes
  Write-StringToStream $bodyStream ("--$boundary--$LF")

  $bodyBytes = $bodyStream.ToArray()
  $uri = "$ApiRoot/imports/$ImportDefinitionGuid/runs"
  $httpClient = New-Object System.Net.Http.HttpClient
  $httpClient.Timeout = [TimeSpan]::FromMinutes(15)
  $httpClient.DefaultRequestHeaders.Add('arena_session_id', $Headers['arena_session_id'])
  $reqContent = New-Object System.Net.Http.ByteArrayContent (,$bodyBytes)
  $reqContent.Headers.ContentType =
    [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("multipart/form-data; boundary=$boundary")
  $resp = $httpClient.PostAsync($uri, $reqContent).GetAwaiter().GetResult()
  $respBody = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
  if (-not $resp.IsSuccessStatusCode) {
    throw "Arena POST $uri failed ($([int]$resp.StatusCode)): $respBody"
  }
  if ([string]::IsNullOrWhiteSpace($respBody)) { throw 'Import run response was empty.' }
  return $respBody | ConvertFrom-Json
}

function Wait-ArenaImportRunComplete {
  param(
    [string]$ApiRoot,
    [hashtable]$Headers,
    [string]$ImportDefinitionGuid,
    [string]$RunGuid,
    [int]$MaxWaitSeconds = 600
  )
  $deadline = (Get-Date).AddSeconds($MaxWaitSeconds)
  $lastStatus = ''
  while ((Get-Date) -lt $deadline) {
    $run = Invoke-ArenaJson -Uri "$ApiRoot/imports/$ImportDefinitionGuid/runs/$RunGuid" -Headers $Headers
    $status = ("$($run.status)").Trim().ToUpperInvariant()
    if ($status -ne $lastStatus) {
      Write-Host ("  Import status: {0} (success={1}, errors={2})" -f $status, $run.successCount, $run.errorCount)
      $lastStatus = $status
    }
    if ($status -in @('COMPLETE', 'ERROR')) { return $run }
    Start-Sleep -Seconds 2
  }
  throw "Import run did not finish within $MaxWaitSeconds seconds (last status: $lastStatus)."
}

function Split-BomCsvDesignatorTokens {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
  return @(($Text -split '[,;]') | ForEach-Object { $_.Trim().Trim('"') } | Where-Object { $_ })
}

function Test-DesignatorsAreOnlyMhOrTp {
  param($Tokens)
  $tok = @($Tokens | Where-Object { $_ })
  if ($tok.Count -eq 0) { return $false }
  foreach ($t in $tok) {
    $ts = ("$t").Trim()
    if ($ts -notmatch '^(?i)MH\d+$' -and $ts -notmatch '^(?i)TP\d+$') {
      return $false
    }
  }
  return $true
}

function Find-BomCsvHeaderLineIndex {
  param($Lines)
  $arr = @($Lines)
  for ($i = 0; $i -lt $arr.Count; $i++) {
    $l = "$($arr[$i])"
    if ($l -match '(?i)Designator' -and $l -match '(?i)Quantity|Qty') { return $i }
  }
  for ($i = 0; $i -lt $arr.Count; $i++) {
    if ("$($arr[$i])" -match '(?i)Designator') { return $i }
  }
  return -1
}

function Get-BomCsvRowField {
  param(
    [Parameter(Mandatory)]$Row,
    [string]$ColumnName
  )
  if ([string]::IsNullOrWhiteSpace($ColumnName)) { return '' }
  foreach ($prop in $Row.PSObject.Properties) {
    if ($prop.Name -eq $ColumnName) {
      if ($null -eq $prop.Value) { return '' }
      return ("" + $prop.Value).Trim()
    }
  }
  return ''
}

function Find-BomCsvColumnName {
  param(
    $ColumnNames,
    $Candidates
  )
  $names = @($ColumnNames | ForEach-Object { "$_" })
  $cands = @($Candidates | ForEach-Object { "$_" })
  foreach ($c in $cands) {
    $want = $c.Trim().ToLowerInvariant()
    foreach ($name in $names) {
      if ("$name".Trim().ToLowerInvariant() -eq $want) { return "$name" }
    }
  }
  return $null
}

function Test-ShouldSkipAltiumBomCsvRow {
  param(
    [string]$PartNumber,
    [string]$RefDes,
    [string]$AssemblyPartNumber,
    [string]$Level
  )
  $pn = ("" + $PartNumber).Trim()
  if ($pn -imatch '(?i)^01-PJM\s+Part[_ ]?Number$') { return $true }
  if ($pn -imatch '(?i)^(WDTS\s*)?Part[_ ]?Number$') { return $true }
  if ($pn -imatch '(?i)^ITEM\s*NUMBER$') { return $true }
  if (-not [string]::IsNullOrWhiteSpace($AssemblyPartNumber) -and ($pn -ieq $AssemblyPartNumber.Trim())) {
    return $true
  }
  if (-not $pn) { return $true }
  if ($Level -and ("$Level".Trim() -ne '1')) { return $true }
  $tokens = Split-BomCsvDesignatorTokens -Text $RefDes
  if (Test-DesignatorsAreOnlyMhOrTp -Tokens $tokens) { return $true }
  return $false
}

function Get-AltiumBomLinesFromCsv {
  param(
    [Parameter(Mandatory)][string]$CsvPath,
    [Parameter(Mandatory)][string]$AssemblyPartNumber
  )
  $lines = @([System.IO.File]::ReadAllLines($CsvPath, [System.Text.Encoding]::UTF8))
  $headerIdx = Find-BomCsvHeaderLineIndex -Lines $lines
  if ($headerIdx -lt 0) {
    throw 'Altium BOM CSV has no header row with Designator (and Quantity/Qty).'
  }
  $csvBlock = ($lines[$headerIdx..($lines.Count - 1)]) -join [Environment]::NewLine
  $rows = @($csvBlock | ConvertFrom-Csv)
  if (-not $rows -or $rows.Count -eq 0) {
    throw 'No data rows found in Altium BOM CSV after the header row.'
  }
  $cols = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
  $colPn = Find-BomCsvColumnName -ColumnNames $cols -Candidates @(
    'item number', 'part number', 'internal part number', '01-pjm part_number'
  )
  $colQty = Find-BomCsvColumnName -ColumnNames $cols -Candidates @('quantity', 'qty')
  $colDes = Find-BomCsvColumnName -ColumnNames $cols -Candidates @(
    'reference designator', 'reference designators', 'designator', 'designators'
  )
  $colLevel = Find-BomCsvColumnName -ColumnNames $cols -Candidates @('level')
  if (-not $colPn) {
    throw ("CSV header is missing a part-number column. Columns: {0}" -f ($cols -join ', '))
  }
  if (-not $colQty) {
    throw ("CSV header is missing quantity. Columns: {0}" -f ($cols -join ', '))
  }

  $out = New-Object System.Collections.Generic.List[object]
  foreach ($row in @($rows)) {
    $pn = Get-BomCsvRowField -Row $row -ColumnName $colPn
    $qty = 0
    [void][int]::TryParse((Get-BomCsvRowField -Row $row -ColumnName $colQty), [ref]$qty)
    $refDes = if ($colDes) { Get-BomCsvRowField -Row $row -ColumnName $colDes } else { '' }
    $level = if ($colLevel) { Get-BomCsvRowField -Row $row -ColumnName $colLevel } else { '1' }
    if (Test-ShouldSkipAltiumBomCsvRow -PartNumber $pn -RefDes $refDes -AssemblyPartNumber $AssemblyPartNumber -Level $level) {
      continue
    }
    [void]$out.Add([pscustomobject]@{
      PartNumber = $pn
      Quantity   = $qty
      RefDes     = $refDes
    })
  }
  if ($out.Count -eq 0) {
    throw 'No BOM lines remain after skipping headers, MH#/TP#, and assembly row.'
  }
  return ,$out.ToArray()
}

function Get-ArenaAssemblyItemGuid {
  param(
    [string]$ApiRoot,
    [hashtable]$Headers,
    [string]$PartNumber
  )
  $pnEnc = [Uri]::EscapeDataString($PartNumber.Trim())
  $lookup = Invoke-ArenaJson -Uri "$ApiRoot/items?number=$pnEnc&limit=1" -Headers $Headers
  if (-not $lookup.results -or $lookup.results.Count -eq 0) {
    throw "No Arena item found for assembly part number '$($PartNumber.Trim())'."
  }
  return "$($lookup.results[0].guid)".Trim()
}

function Clear-ArenaItemBom {
  param(
    [string]$ApiRoot,
    [hashtable]$Headers,
    [string]$AssemblyItemGuid
  )
  $bom = Invoke-ArenaJson -Uri "$ApiRoot/items/$AssemblyItemGuid/bom" -Headers $Headers
  $n = 0
  if (-not $bom -or -not $bom.results) { return 0 }
  foreach ($line in @($bom.results)) {
    if (-not $line.guid) { continue }
    $lg = "$($line.guid)".Trim()
    [void](Invoke-ArenaJson -Uri "$ApiRoot/items/$AssemblyItemGuid/bom/$lg" -Method Delete -Headers $Headers)
    $n++
  }
  return $n
}

function Import-ArenaBomFromCsvViaRestApi {
  param(
    [string]$ApiRoot,
    [hashtable]$Headers,
    [string]$AssemblyPartNumber,
    [string]$CsvPath,
    [string]$Mode
  )
  $assemblyGuid = Get-ArenaAssemblyItemGuid -ApiRoot $ApiRoot -Headers $Headers -PartNumber $AssemblyPartNumber
  $bomLines = Get-AltiumBomLinesFromCsv -CsvPath $CsvPath -AssemblyPartNumber $AssemblyPartNumber
  Write-Host ("  Assembly item GUID : {0}" -f $assemblyGuid)
  Write-Host ("  BOM lines to import: {0}" -f $bomLines.Count)

  $modeKey = ("$Mode").Trim().ToUpperInvariant()
  if ($modeKey -eq 'REPLACE') {
    $removed = Clear-ArenaItemBom -ApiRoot $ApiRoot -Headers $Headers -AssemblyItemGuid $assemblyGuid
    Write-Host ("  REPLACE: removed {0} existing BOM line(s)." -f $removed)
  }

  $childCache = @{}
  $added = 0
  $failed = New-Object System.Collections.Generic.List[string]
  foreach ($bl in $bomLines) {
    $pn = $bl.PartNumber
    if (-not $childCache.ContainsKey($pn)) {
      $pnEnc = [Uri]::EscapeDataString($pn)
      $lookup = Invoke-ArenaJson -Uri "$ApiRoot/items?number=$pnEnc&limit=1" -Headers $Headers
      if (-not $lookup.results -or $lookup.results.Count -eq 0) {
        $failed.Add("No Arena item for $pn")
        continue
      }
      $childCache[$pn] = "$($lookup.results[0].guid)".Trim()
    }
    $lineBody = @{
      refDes   = $bl.RefDes
      quantity = $bl.Quantity
      item     = @{ guid = $childCache[$pn] }
    } | ConvertTo-Json -Compress
    try {
      [void](Invoke-ArenaJson -Uri "$ApiRoot/items/$assemblyGuid/bom" -Method Post -Headers $Headers -Body $lineBody)
      $added++
    } catch {
      $arenaErr = @($_.Exception.Message -split '[\r\n]' | Where-Object {
        $_ -match 'Arena error \d+:|Arena response body:'
      })
      if ($arenaErr.Count -gt 0) {
        $failed.Add("$pn : $($arenaErr -join ' | ')")
      } else {
        $failed.Add("$pn : $(($_.Exception.Message -split '[\r\n]')[0].Trim())")
      }
      Write-Host ("  BOM line POST failed for {0}:" -f $pn) -ForegroundColor Yellow
      foreach ($line in ($_.Exception.Message -split '[\r\n]')) {
        if ($line.Trim()) { Write-Host "    $line" }
      }
    }
  }

  Write-Host ("  Added {0} BOM line(s) via REST API." -f $added)
  if ($failed.Count -gt 0) {
    Write-Host ("  Failed {0} line(s):" -f $failed.Count) -ForegroundColor Red
    foreach ($f in $failed) { Write-Host "    $f" }
    throw 'One or more BOM lines could not be added.'
  }
}

function Resolve-ArenaImportMode {
  param(
    [string]$RequestedMode,
    [int]$ArenaBomLineCount
  )
  $key = ("$RequestedMode").Trim().ToLowerInvariant()
  if ($key -ne 'auto') {
    return $RequestedMode.Trim().ToUpperInvariant()
  }
  if ($ArenaBomLineCount -eq 0) { return 'CREATE' }
  return 'REPLACE'
}

function Write-ArenaItemBomImportDefinitionCatalog {
  param(
    [string]$ApiRoot,
    [hashtable]$Headers,
    [string]$RequestedMode
  )
  try {
    $all = Invoke-ArenaJson -Uri "$ApiRoot/imports?resource=ITEM_BOM&limit=50" -Headers $Headers
    if (-not $all.results -or $all.results.Count -eq 0) {
      Write-Host '  (No ITEM_BOM import definitions exist in this workspace.)'
      return
    }
    Write-Host '  Available ITEM_BOM import definitions in this workspace:'
    foreach ($d in @($all.results)) {
      Write-Host ("    name={0}  mode={1}  number={2}  guid={3}" -f $d.name, $d.mode, $d.number, $d.guid)
    }
    if ($RequestedMode) {
      $match = @($all.results | Where-Object { ("$($_.mode)").ToUpperInvariant() -eq $RequestedMode })
      if ($match.Count -eq 0) {
        Write-Host ("  None with mode={0}. Create one in Arena Import Engine for ITEM_BOM." -f $RequestedMode)
      }
    }
  } catch {
    Write-Host ("  (Could not list import definitions: {0})" -f $_.Exception.Message)
  }
}

function Resolve-ArenaBomImportDefinition {
  param(
    [string]$ApiRoot,
    [hashtable]$Headers,
    [string]$PreferredGuid,
    [Parameter(Mandatory = $true)] [string]$Mode
  )
  if (-not [string]::IsNullOrWhiteSpace($PreferredGuid)) {
    Write-Host ("  Using import definition GUID from parameter/env: {0}" -f $PreferredGuid.Trim())
    return $PreferredGuid.Trim()
  }
  $modeEnc = [Uri]::EscapeDataString($Mode)
  $defs = Invoke-ArenaJson -Uri "$ApiRoot/imports?resource=ITEM_BOM&mode=$modeEnc&limit=50" -Headers $Headers
  if (-not $defs.results -or $defs.results.Count -eq 0) {
    Write-Host ''
    Write-Host ("No Arena Import Definition found for resource=ITEM_BOM mode=$Mode.") -ForegroundColor Red
    Write-ArenaItemBomImportDefinitionCatalog -ApiRoot $ApiRoot -Headers $Headers -RequestedMode $Mode
    Write-Host ''
    Write-Host 'To fix this, either:'
    Write-Host ("  1. In Arena, create an Import Definition: resource=Item BOM, mode=$Mode,")
    Write-Host '     with column mapping that matches your Altium CSV (item number, quantity, ref des, etc.).'
    Write-Host '  2. Copy its GUID into Script Launcher "Import definition GUID" or set env:'
    Write-Host '     ARENA_BOM_IMPORT_DEFINITION_GUID=<guid>'
    Write-Host '  3. For an empty working-rev BOM (like 296-15-001), use mode CREATE, not REPLACE.'
    Write-Host '  Or leave GUID blank: the script will add BOM lines via REST API instead.'
    return $null
  }
  $pick = $defs.results | Select-Object -First 1
  if ($defs.results.Count -gt 1) {
    Write-Host ('  Note: {0} definitions match mode={1}; using first: {2}' -f $defs.results.Count, $Mode, $pick.name)
  }
  Write-Host ("  Using import definition: {0} (mode={1}, number={2}) guid={3}" -f $pick.name, $pick.mode, $pick.number, $pick.guid)
  return "$($pick.guid)".Trim()
}

# ---------- resolve Altium production BOM folder ----------

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
$resolved = $jm.Groups['j'].Value | ConvertFrom-Json
$prodDoc = $resolved.prodDocDir
$targetDir = $resolved.targetDir
Write-Host ("  Production folder : {0}" -f $prodDoc)
Write-Host ("  Target/rev folder : {0}" -f $targetDir)

$searchDirs = @(
  (Join-Path $prodDoc 'BOM'),
  $prodDoc,
  $targetDir
) | Where-Object { Test-Path -LiteralPath $_ -PathType Container }

$csvCandidates = @()
$xlsxCandidates = @()
$csvUsedDir = $null
$xlsxUsedDir = $null
foreach ($d in $searchDirs) {
  if ($csvCandidates.Count -eq 0) {
    $csvCandidates = @(Get-ChildItem -LiteralPath $d -File -ErrorAction SilentlyContinue |
      Where-Object { $_.Extension -ieq '.csv' })
    if ($csvCandidates.Count -gt 0) { $csvUsedDir = $d }
  }
  if ($xlsxCandidates.Count -eq 0) {
    $xlsxCandidates = @(Get-ChildItem -LiteralPath $d -File -ErrorAction SilentlyContinue |
      Where-Object { $_.Extension -ieq '.xlsx' -or $_.Extension -ieq '.xls' })
    if ($xlsxCandidates.Count -gt 0) { $xlsxUsedDir = $d }
  }
}

if ($csvCandidates.Count -eq 0) {
  throw ("No .csv file found under production BOM search paths:`n  - " + ($searchDirs -join "`n  - "))
}

$pickedCsv = $csvCandidates | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
$pickedXlsx = $null
if ($xlsxCandidates.Count -gt 0) {
  $pickedXlsx = $xlsxCandidates | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
}

Write-Host ''
Write-Host '=== BOM files ==='
Write-Host ("  CSV search folder : {0}" -f $csvUsedDir)
Write-Host ("  CSV for upload    : {0}" -f $pickedCsv.FullName)
Write-Host ("  CSV last modified : {0}" -f $pickedCsv.LastWriteTime)
if ($pickedXlsx) {
  Write-Host ("  Excel for headers : {0}" -f $pickedXlsx.FullName)
  Write-Host ("  Excel modified    : {0}" -f $pickedXlsx.LastWriteTime)
} else {
  Write-Host '  Excel for headers : (none - using CSV lines 2-3)'
}

# ---------- PCBA / PCB part numbers ----------

Write-Host ''
Write-Host '=== Part numbers from Excel export ==='
$pcbaPn = $null
$pcbPn = $null
if ($pickedXlsx) {
  $hdr = Get-PcbaPcbFromExcelBom -ExcelPath $pickedXlsx.FullName
  $pcbaPn = $hdr.Pcba
  $pcbPn = $hdr.Pcb
}
if (-not $pcbaPn -or -not $pcbPn) {
  $csvLines = @(Get-Content -LiteralPath $pickedCsv.FullName -ErrorAction Stop)
  if ($csvLines.Count -ge 2 -and -not $pcbaPn) {
    $pcbaPn = Get-PartNumberFromBomText -Text $csvLines[1] -MiddleSegmentPattern '10|15'
  }
  if ($csvLines.Count -ge 3 -and -not $pcbPn) {
    $pcbPn = Get-PartNumberFromBomText -Text $csvLines[2] -MiddleSegmentPattern '20|25'
  }
}
if ($pcbaPn) {
  Write-Host ("  PCBA part number : {0}" -f $pcbaPn)
} else {
  Write-Host '  PCBA part number : (not found - expected PPP-10-NN or PPP-15-NN in Excel row 2 / CSV line 2)' -ForegroundColor Yellow
}
if ($pcbPn) {
  Write-Host ("  PCB part number  : {0}" -f $pcbPn)
} else {
  Write-Host '  PCB part number  : (not found - expected PPP-20-NN or PPP-25-NN in Excel row 3 / CSV line 3)' -ForegroundColor Yellow
}

# ---------- Arena BOM compare (v2) or empty message ----------

Write-Host ''
Write-Host "=== Compare Arena vs Altium BOM v2 ==="
$arenaBomLines = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File $getBomScript -PartNumber $PartNumber -Format csv 2>&1
$arenaBomExit = $LASTEXITCODE
$arenaRowCount = 0
if ($arenaBomExit -eq 0 -and $arenaBomLines) {
  try {
    $parsed = ($arenaBomLines -join "`n") | ConvertFrom-Csv
    if ($parsed) { $arenaRowCount = @($parsed).Count }
  } catch {
    $arenaRowCount = 0
  }
}

$resolvedImportMode = Resolve-ArenaImportMode -RequestedMode $ImportMode -ArenaBomLineCount $arenaRowCount

Write-Host ''
Write-Host '=== Planned Arena BOM import (after you confirm) ==='
Write-Host ("  Import mode     : {0}" -f $resolvedImportMode)
if ($ImportMode.Trim().ToLowerInvariant() -eq 'auto') {
  if ($arenaRowCount -eq 0) {
    Write-Host '  (Auto chose CREATE because the Arena BOM is empty.)'
  } else {
    Write-Host '  (Auto chose REPLACE because the Arena BOM already has lines.)'
  }
}
if (-not [string]::IsNullOrWhiteSpace($ImportDefinitionGuid)) {
  Write-Host ("  Definition GUID : {0} (from parameter/env)" -f $ImportDefinitionGuid.Trim())
} else {
  Write-Host ("  Definition GUID : (search ITEM_BOM + mode {0}; else REST API line import)" -f $resolvedImportMode)
}
if ($ForceDirectApi) {
  Write-Host '  Import method   : REST API only (-ForceDirectApi)'
}

if ($arenaBomExit -ne 0 -or $arenaRowCount -eq 0) {
  Write-Host ("Arena BOM is currently empty for this part ($PartNumber).") -ForegroundColor DarkGray
  Write-Host '  (Skipping Compare Arena vs Altium BOM v2.)'
} else {
  Write-Host ("  Arena BOM has {0} line(s). Running comparison..." -f $arenaRowCount)
  Write-Host ''
  & powershell.exe -NoProfile -ExecutionPolicy Bypass `
    -File $compareAutoScript -PartNumber $PartNumber -Sheet '1'
  $compareExit = $LASTEXITCODE
  Write-Host ''
  if ($compareExit -eq 0) {
    Write-Host '  Compare finished with no description/PJM mismatches.' -ForegroundColor Green
  } else {
    Write-Host ("  Compare finished with exit code {0} (review differences above)." -f $compareExit) -ForegroundColor Yellow
  }
}

# ---------- preview stop or upload ----------

if (-not $ConfirmUpload) {
  Write-Host ''
  Write-Host '=== Ready to upload ==='
  Write-Host 'Review the information above, then click Confirm Upload in the launcher.'
  Write-Host '---UPLOAD-CONFIRM-PROMPT---'
  [pscustomobject]@{
    confirmUpload = $true
    message       = 'Upload BOM CSV to Arena? Review the report above, then confirm.'
    partNumber    = $PartNumber.Trim()
    pcba          = $pcbaPn
    pcb           = $pcbPn
    csvPath       = $pickedCsv.FullName
    csvName       = $pickedCsv.Name
  } | ConvertTo-Json -Compress
  exit 12
}

Write-Host ''
Write-Host '=== Uploading BOM to Arena ==='

$apiRoot = $BaseUrl.TrimEnd('/')
$loginBody = @{
  email       = $Email
  password    = $env:ARENA_PASSWORD
  workspaceId = [long]$WorkspaceId
} | ConvertTo-Json
$login = Invoke-ArenaJson -Uri "$apiRoot/login" -Method Post -Body $loginBody
$sid = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sid)) { $sid = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sid)) { throw 'Login response did not include arenaSessionId.' }
$headers = @{
  arena_session_id = $sid
  'Content-Type'   = 'application/json'
}

Write-Host ("  CSV file: {0}" -f $pickedCsv.FullName)

$importGuid = $null
if (-not $ForceDirectApi) {
  $importGuid = Resolve-ArenaBomImportDefinition -ApiRoot $apiRoot -Headers $headers `
    -PreferredGuid $ImportDefinitionGuid -Mode $resolvedImportMode
}

if ($importGuid) {
  Write-Host ("=== Uploading via Arena Import Engine (ITEM_BOM {0}) ===" -f $resolvedImportMode)
  $run = Invoke-ArenaMultipartImportRun -ApiRoot $apiRoot -Headers $headers `
    -ImportDefinitionGuid $importGuid -CsvPath $pickedCsv.FullName -Commit
  $runGuid = "$($run.guid)".Trim()
  if (-not $runGuid) { throw 'Import run did not return a guid.' }
  Write-Host ("  Import run started: {0}" -f $runGuid)

  $final = Wait-ArenaImportRunComplete -ApiRoot $apiRoot -Headers $headers `
    -ImportDefinitionGuid $importGuid -RunGuid $runGuid
  $finalStatus = ("$($final.status)").Trim().ToUpperInvariant()
  if ($finalStatus -eq 'ERROR' -or ([int]$final.errorCount -gt 0)) {
    Write-Host ("Import completed with errors (status={0}, errors={1})." -f $finalStatus, $final.errorCount) -ForegroundColor Red
    Write-Host 'Download the error content file from Arena Import Runs for details.'
    exit 1
  }
} else {
  Write-Host ("=== Uploading via Arena REST API (ITEM_BOM {0}, no Import Definition) ===" -f $resolvedImportMode)
  Import-ArenaBomFromCsvViaRestApi -ApiRoot $apiRoot -Headers $headers `
    -AssemblyPartNumber $PartNumber.Trim() -CsvPath $pickedCsv.FullName -Mode $resolvedImportMode
}

Write-Host ''
Write-Host ("Done. BOM import committed for {0} ({1})." -f $PartNumber.Trim(), $pickedCsv.Name) -ForegroundColor Green
exit 0
