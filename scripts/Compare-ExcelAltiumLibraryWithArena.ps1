<#
.SYNOPSIS
  Compare an Excel library workbook (all sheets with 01-PJM Part_Number) against Arena item names.

.DESCRIPTION
  For each worksheet in workbook order, if the sheet has the standard PJM header row
  (01-PJM Part_Number plus 03-Description General or the same fallbacks as
  Check-AndUploadAltiumBomExport), each part number is looked up in Arena via the items API
  and the item ``description`` (if present, else ``name``) is compared to the Excel description using
  Test-BomDescriptionsCompatible from Bom-ArenaAltiumDescriptionChecks.ps1.

  Writes ``<basename>_ArenaDiscrepancies.xlsx`` next to the input (override with -OutputPath).
  Output workbook has the same number and order of sheets as the source; each sheet contains
  only columns ``Part Number`` and ``Discrepancy`` (header row plus one row per part that failed).
  Sheets without PJM headers still appear with headers only.

  Credentials: ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID.

.PARAMETER ExcelPath
  Full path to the source .xlsx or .xls library file.

.PARAMETER OutputPath
  Optional full path for the report .xlsx. If omitted, the file is written to your Desktop as
  ``<inputBasename>_ArenaDiscrepancies_yyyyMMdd_HHmmss.xlsx`` (avoids writing next to a temp copy
  when the workbook path is under %TEMP%).
#>
param(
  [Parameter(Mandatory = $true)] [string]$ExcelPath,
  [Parameter(Mandatory = $false)] [string]$OutputPath = ''
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$checksLib = Join-Path $scriptsDir 'Bom-ArenaAltiumDescriptionChecks.ps1'
if (-not (Test-Path -LiteralPath $checksLib)) { throw "Required library missing: $checksLib" }
. $checksLib

if ([string]::IsNullOrWhiteSpace($env:ARENA_EMAIL)) { throw 'Set ARENA_EMAIL.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_PASSWORD)) { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_WORKSPACE_ID)) {
  throw 'Set ARENA_WORKSPACE_ID (launcher workspace dropdown sets this per run).'
}

$srcPath = (Resolve-Path -LiteralPath $ExcelPath).Path
if (-not (Test-Path -LiteralPath $srcPath -PathType Leaf)) { throw "Not a file: $srcPath" }

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
  $desk = [Environment]::GetFolderPath('Desktop')
  if ([string]::IsNullOrWhiteSpace($desk)) {
    $desk = [Environment]::GetFolderPath('UserProfile')
  }
  $stem = [System.IO.Path]::GetFileNameWithoutExtension($srcPath)
  if ([string]::IsNullOrWhiteSpace($stem)) { $stem = 'Library' }
  $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
  $OutputPath = Join-Path $desk ("{0}_ArenaDiscrepancies_{1}.xlsx" -f $stem, $stamp)
}
else {
  $OutputPath = $OutputPath.Trim()
  $outDir = Split-Path -LiteralPath $OutputPath -Parent
  if (-not [string]::IsNullOrWhiteSpace($outDir) -and -not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  }
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

function Sanitize-ExcelWorksheetName([string]$name) {
  $n = ($name -replace '[\\/*?:\[\]]', '_').Trim()
  if ([string]::IsNullOrWhiteSpace($n)) { $n = 'Sheet' }
  if ($n.Length -gt 31) { $n = $n.Substring(0, 31) }
  return $n
}

function Set-OutSheetNameSafely {
  param($Worksheet, [string]$DesiredName)
  $base = Sanitize-ExcelWorksheetName $DesiredName
  $name = $base
  $suffix = 2
  while ($true) {
    try {
      $Worksheet.Name = $name
      return
    } catch {
      $suffixStr = " ($suffix)"
      $max = 31 - $suffixStr.Length
      if ($max -lt 1) { $max = 1 }
      $trunc = if ($base.Length -le $max) { $base } else { $base.Substring(0, $max) }
      $name = $trunc + $suffixStr
      $suffix++
      if ($suffix -gt 99) { $Worksheet.Name = 'Sheet'; return }
    }
  }
}

function New-ExcelWorkbookWithSheetCount {
  <#
  Build a workbook with exactly $SheetCount worksheets. Excel COM in PowerShell often fails on
  Sheets.Add($null, ...) — use [System.Reflection.Missing]::Value for the Before parameter, and
  pass an explicit worksheet type (-4167 = xlWorksheet) when supported.
  #>
  param(
    [Parameter(Mandatory = $true)] $ExcelApp,
    [Parameter(Mandatory = $true)] [int]$SheetCount
  )
  if ($SheetCount -lt 1) { throw "SheetCount must be >= 1, got $SheetCount." }
  $mis = [System.Reflection.Missing]::Value
  $xlWorksheet = -4167
  $wb = $ExcelApp.Workbooks.Add()
  try {
    while ([int]$wb.Sheets.Count -gt $SheetCount) {
      $wb.Sheets.Item([int]$wb.Sheets.Count).Delete()
    }
    while ([int]$wb.Sheets.Count -lt $SheetCount) {
      $last = $wb.Sheets.Item([int]$wb.Sheets.Count)
      $addErr = $null
      try {
        [void]$wb.Sheets.Add($mis, $last, 1, $xlWorksheet)
      } catch {
        $addErr = $_
        try {
          [void]$wb.Worksheets.Add($mis, $last)
          $addErr = $null
        } catch {
          $addErr = $_
        }
      }
      if ($null -ne $addErr) {
        throw "Could not add a worksheet (have $($wb.Sheets.Count), need $SheetCount): $($addErr.Exception.Message)"
      }
    }
  } catch {
    try { $wb.Close($false) } catch {}
    throw
  }
  return $wb
}

$BaseUrl = 'https://api.arenasolutions.com/v1'
$loginBody = @{
  email       = $env:ARENA_EMAIL
  password    = $env:ARENA_PASSWORD
  workspaceId = [long]$env:ARENA_WORKSPACE_ID
} | ConvertTo-Json

Write-Host '=== Arena login ==='
try {
  $login = Invoke-ArenaJson -Uri "$BaseUrl/login" -Method Post -Body $loginBody
} catch {
  throw "Arena login failed: $($_.Exception.Message)"
}
$sessionId = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sessionId)) { $sessionId = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sessionId)) { throw 'Login response did not include arenaSessionId.' }

$arenaHeaders = @{
  arena_session_id = $sessionId
  'Content-Type'   = 'application/json'
}

# Cache key -> Arena side text (description else name), $null = no item, error token = HTTP failure
$arenaNameCache = @{}
$ArenaLookupErrorToken = '__ARENA_LOOKUP_HTTP_ERROR__'

function Get-ArenaItemSideTextCached {
  param([string]$PartNumber)
  $key = $PartNumber.Trim()
  if ($arenaNameCache.ContainsKey($key)) { return $arenaNameCache[$key] }
  $pnEnc = [Uri]::EscapeDataString($key)
  $viewEnc = [Uri]::EscapeDataString('full')
  $uri = "$BaseUrl/items?number=$pnEnc&limit=50&responseview=$viewEnc"
  try {
    $items = Invoke-ArenaJson -Uri $uri -Headers $arenaHeaders
  } catch {
    $arenaNameCache[$key] = $ArenaLookupErrorToken
    return $ArenaLookupErrorToken
  }
  if (-not $items.results -or $items.results.Count -eq 0) {
    $arenaNameCache[$key] = $null
    return $null
  }
  $it = $items.results[0]
  $desc = if ($null -ne $it.description) { ("$($it.description)").Trim() } else { '' }
  $nm = if ($null -ne $it.name) { ("$($it.name)").Trim() } else { '' }
  $side = if ($desc) { $desc } else { $nm }
  $arenaNameCache[$key] = $side
  return $side
}

function Format-DiscrepancyOneLine {
  param([string]$Message, [string]$ArenaName, [string]$ExcelDesc)
  $m = ($Message -replace "`r`n", ' ' -replace "`n", ' ').Trim()
  $a = ($ArenaName -replace "`r`n", ' ' -replace "`n", ' ').Trim()
  $e = ($ExcelDesc -replace "`r`n", ' ' -replace "`n", ' ').Trim()
  if ($a.Length -gt 220) { $a = $a.Substring(0, 217) + '...' }
  if ($e.Length -gt 220) { $e = $e.Substring(0, 217) + '...' }
  return ("{0} | Arena: {1} | Excel (03-Description General): {2}" -f $m, $a, $e).Trim()
}

Write-Host ''
Write-Host ("=== Reading library workbook: {0} ===" -f $srcPath)

$excel = $null
$srcWb = $null
$outWb = $null

try {
  $excel = New-Object -ComObject Excel.Application
} catch {
  throw "Could not start Excel COM: $($_.Exception.Message)"
}

$sheetReports = New-Object System.Collections.Generic.List[object]

try {
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $srcWb = $excel.Workbooks.Open($srcPath, 0, $true)

  $sheetCount = [int]$srcWb.Sheets.Count
  Write-Host ("  Sheets: {0}" -f $sheetCount)

  for ($si = 1; $si -le $sheetCount; $si++) {
    $sws = $srcWb.Sheets.Item($si)
    $sn = [string]$sws.Name
    Write-Host ("  Scanning sheet {0}/{1}: {2}" -f $si, $sheetCount, $sn)
    $disc = New-Object System.Collections.Generic.List[hashtable]
    $hasPjm = $false
    try {
      $pj = Get-PjmPartRowsFromWorksheet -Worksheet $sws
      $hasPjm = $true
    } catch {
      $sheetReports.Add([pscustomobject]@{
          SourceName = $sn
          HasPjm     = $false
          DiscRows   = @()
        })
      continue
    }

    foreach ($pr in $pj.Rows) {
      $pjn = $pr.PartNumber
      if (-not $pjn) { continue }
      $arenaName = Get-ArenaItemSideTextCached -PartNumber $pjn
      if ($arenaName -eq $ArenaLookupErrorToken) {
        [void]$disc.Add(@{
            Part = $pjn
            Text = "Arena API error while looking up part number."
          })
        continue
      }
      if ($null -eq $arenaName) {
        [void]$disc.Add(@{
            Part = $pjn
            Text = 'Arena has no item with this part number.'
          })
        continue
      }
      $chk = Test-BomDescriptionsCompatible -ArenaDescription $arenaName -AltiumDescription $pr.Description -PartNumber $pjn
      if (-not $chk.Ok) {
        [void]$disc.Add(@{
            Part = $pjn
            Text = (Format-DiscrepancyOneLine -Message $chk.Message -ArenaName $arenaName -ExcelDesc $pr.Description)
          })
      }
    }

    $sheetReports.Add([pscustomobject]@{
        SourceName = $sn
        HasPjm     = $true
        DiscRows   = @($disc)
      })
  }

  $srcWb.Close($false)
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($srcWb)
  $srcWb = $null

  Write-Host ''
  Write-Host ("=== Writing report: {0} ===" -f $OutputPath)

  $need = $sheetCount
  $outWb = New-ExcelWorkbookWithSheetCount -ExcelApp $excel -SheetCount $need

  for ($i = 1; $i -le $need; $i++) {
    $ows = $outWb.Sheets.Item($i)
    $rep = $sheetReports[$i - 1]
    Set-OutSheetNameSafely -Worksheet $ows -DesiredName $rep.SourceName
    $ows.Cells.Item(1, 1).Value2 = 'Part Number'
    $ows.Cells.Item(1, 2).Value2 = 'Discrepancy'
    $r = 2
    foreach ($d in $rep.DiscRows) {
      $ows.Cells.Item($r, 1).Value2 = $d.Part
      $ows.Cells.Item($r, 2).Value2 = $d.Text
      $r++
    }
  }

  if (Test-Path -LiteralPath $OutputPath) {
    Remove-Item -LiteralPath $OutputPath -Force -ErrorAction Stop
  }
  $outWb.SaveAs($OutputPath, 51)
  $outWb.Close($true)
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($outWb)
  $outWb = $null

  $totalDisc = 0
  foreach ($rep in $sheetReports) { $totalDisc += $rep.DiscRows.Count }
  Write-Host ''
  Write-Host ("Done. Total discrepancy rows: {0}. Output:`n  {1}" -f $totalDisc, $OutputPath) -ForegroundColor DarkGreen
}
finally {
  if ($srcWb) {
    try { $srcWb.Close($false) } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($srcWb) } catch {}
  }
  if ($outWb) {
    try { $outWb.Close($false) } catch {}
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($outWb) } catch {}
  }
  if ($excel) {
    try { $excel.Quit() } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
  }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

exit 0
