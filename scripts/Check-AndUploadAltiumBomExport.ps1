<#
.SYNOPSIS
  Check and Create CSV Altium BOM export: validate PJM sheet vs Arena and Name vs description, then run Convert BOM (Excel -> CSV) v2.

.DESCRIPTION
  1. Resolves the Altium production folder (Open-AltiumPartFolder.ps1 -Open ResolveJson).
  2. Picks the most recently modified .xlsx / .xls under the BOM folder (same search order
     as Compare-ArenaAltiumBomAuto).
  3. Logs into Arena and, for each 01-PJM Part_Number on the PJM sheet, looks up the **item**
     by number (REST items API). The Arena side of the compare uses the item **description**
     when present, otherwise the item **name** (same spec-aware rules as BOM compare v2:
     voltage, current, footprint, passive-style values; spacing/case tolerant). If the PJM
     **Name** cell matches the Arena **manufacturer part number** (standard field or
     additionalAttributes), that line is accepted without comparing descriptions (Name vs
     03-Description General or Arena vs Excel description). Rows are
     checked even when the assembly has no Arena BOM or the part is not on that BOM.
  4. Opens the workbook in Excel COM, reads the PJM sheet (default index 2), finds header
     "01-PJM Part_Number" and description column "03-Description General" (or 01-PJM ...
     Description / Description as fallbacks). If a "Name" or "Comment" column exists on that same sheet,
     only lines whose 01-PJM Part_Number is 098-12-XXX or 098-13-XXX are checked for
     unreasonable differences between Name and 03-Description General (same rules), unless
     Name matched Arena MPN/item-name as above.
  5. If any row fails (including no Arena item for that part number), throws with details
     (no CSV export).
  6. If all checks pass, runs Convert-AltiumBomAuto.ps1 for the same part number.

  Credentials: ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID.

.PARAMETER PartNumber
  Internal part number (PPP-10-XXX etc.) used for Arena BOM and Altium folder resolution.

.PARAMETER PjmSheet
  1-based index or sheet name for the PJM tab. Default: 2.

.PARAMETER CreateCsv
  When set (default in launcher), run Convert BOM (Excel -> CSV) v2 after checks pass.
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [Parameter(Mandatory = $false)] [string]$PjmSheet = '2',
  [switch]$CreateCsv
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$openScript    = Join-Path $scriptsDir 'Open-AltiumPartFolder.ps1'
$convertScript = Join-Path $scriptsDir 'Convert-AltiumBomAuto.ps1'
$checksLib     = Join-Path $scriptsDir 'Bom-ArenaAltiumDescriptionChecks.ps1'
foreach ($p in @($openScript, $convertScript, $checksLib)) {
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

$script:arenaBomExportLineCache = @{}
$script:arenaHeaders = $null
$ArenaLookupErrorToken = '__ARENA_LOOKUP_HTTP_ERROR__'

function Get-ArenaManufacturerPartNumberFromItem {
  param([object]$Item)
  if (-not $Item) { return '' }
  foreach ($prop in @(
      'manufacturerPartNumber', 'manufacturer_part_number', 'mfgPartNumber', 'mpn',
      'manufacturerPart', 'supplierPartNumber'
    )) {
    try {
      $v = $Item.$prop
      if ($null -ne $v) {
        $t = ("$v").Trim()
        if ($t) { return $t }
      }
    } catch {}
  }
  try {
    $mfg = $Item.manufacturer
    if ($mfg) {
      foreach ($p in @('partNumber', 'part_number', 'number')) {
        $v = $mfg.$p
        if ($null -ne $v) {
          $t = ("$v").Trim()
          if ($t) { return $t }
        }
      }
    }
  } catch {}
  if ($Item.additionalAttributes) {
    foreach ($a in @($Item.additionalAttributes)) {
      $api = ("$($a.apiName)").ToLowerInvariant()
      $an = ("$($a.name)").ToLowerInvariant()
      if ($api -match 'manufacturer|mfg|mpn|supplier.*part' -or $an -match 'manufacturer|mfg|mpn|supplier.*part') {
        $val = ("$($a.value)").Trim()
        if ($val) { return $val }
      }
    }
  }
  return ''
}

function Get-ArenaBomExportLineContextCached {
  <#
  One GET /items per part number: Arena description side text, manufacturer part number, status.
  #>
  param([string]$PartNumber)
  $key = $PartNumber.Trim()
  if ($script:arenaBomExportLineCache.ContainsKey($key)) {
    return $script:arenaBomExportLineCache[$key]
  }
  $pnEnc = [Uri]::EscapeDataString($key)
  $viewEnc = [Uri]::EscapeDataString('full')
  $uri = "https://api.arenasolutions.com/v1/items?number=$pnEnc&limit=50&responseview=$viewEnc"
  $rec = $null
  try {
    $items = Invoke-ArenaJson -Uri $uri -Headers $script:arenaHeaders
  } catch {
    $rec = [pscustomobject]@{
      ApiError = $true
      NoItem   = $false
      SideText = $ArenaLookupErrorToken
      Mpn      = ''
      ItemName = ''
    }
    $script:arenaBomExportLineCache[$key] = $rec
    return $rec
  }
  if (-not $items -or -not $items.results -or $items.results.Count -eq 0) {
    $rec = [pscustomobject]@{
      ApiError = $false
      NoItem   = $true
      SideText = $null
      Mpn      = ''
      ItemName = ''
    }
    $script:arenaBomExportLineCache[$key] = $rec
    return $rec
  }
  $it = $items.results[0]
  $desc = if ($null -ne $it.description) { ("$($it.description)").Trim() } else { '' }
  $nm = if ($null -ne $it.name) { ("$($it.name)").Trim() } else { '' }
  $side = if ($desc) { $desc } else { $nm }
  $mpn = Get-ArenaManufacturerPartNumberFromItem -Item $it
  $rec = [pscustomobject]@{
    ApiError = $false
    NoItem   = $false
    SideText = $side
    Mpn      = $mpn
    ItemName = $nm
  }
  $script:arenaBomExportLineCache[$key] = $rec
  return $rec
}

function Test-ShouldRunNameDescriptionGeneralCheck([string]$PartNumber) {
  if ([string]::IsNullOrWhiteSpace($PartNumber)) { return $false }
  return [bool]($PartNumber.Trim() -match '^(?i)098-(12|13)-\d{3}$')
}

# ---------- Resolve Altium + pick BOM Excel ----------

Write-Host "=== Resolving Altium folder for $PartNumber ==="
$resolveOut = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File $openScript -PartNumber $PartNumber.Trim() -Open ResolveJson 2>&1
if ($LASTEXITCODE -ne 0) {
  throw "Open-AltiumPartFolder.ps1 -Open ResolveJson failed (exit $LASTEXITCODE):`n$($resolveOut -join [Environment]::NewLine)"
}
$joined = ($resolveOut | ForEach-Object { "$_" }) -join "`n"
$rxJson = [regex]'---BEGIN-RESOLVE-JSON---\s*(?<j>\{.*\})\s*---END-RESOLVE-JSON---'
$jm = $rxJson.Match($joined)
if (-not $jm.Success) {
  throw "Could not find resolve JSON markers. Output:`n$joined"
}
$resolved  = $jm.Groups['j'].Value | ConvertFrom-Json
$prodDoc   = $resolved.prodDocDir
$targetDir = $resolved.targetDir
Write-Host ("  Production folder : {0}" -f $prodDoc)
Write-Host ("  Target/rev folder : {0}" -f $targetDir)

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
    break
  }
}
if ($candidates.Count -eq 0) {
  throw ("No .xlsx or .xls file found under any of:`n  - " + ($searchDirs -join "`n  - "))
}

$picked = $candidates | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
Write-Host ''
Write-Host ("=== BOM workbook: {0} ===" -f $picked.FullName)

# ---------- Arena session (item-by-number lookup; no BOM CSV required) ----------

Write-Host ''
Write-Host "=== Arena login (item description vs Excel per 01-PJM Part_Number) ==="
$loginBody = @{
  email       = $env:ARENA_EMAIL
  password    = $env:ARENA_PASSWORD
  workspaceId = [long]$env:ARENA_WORKSPACE_ID
} | ConvertTo-Json
try {
  $login = Invoke-ArenaJson -Uri 'https://api.arenasolutions.com/v1/login' -Method Post -Body $loginBody
} catch {
  throw "Arena login failed: $($_.Exception.Message)"
}
$sessionId = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sessionId)) { $sessionId = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sessionId)) { throw 'Login response did not include arenaSessionId.' }
$script:arenaHeaders = @{
  arena_session_id = $sessionId
  'Content-Type'   = 'application/json'
}
Write-Host '  Logged in; will use item description (else name) for each BOM line part number.'

# ---------- Excel: PJM sheet ----------

$excel = $null
$wb = $null
$tempCopy = $null
$failures = New-Object System.Collections.Generic.List[string]

try {
  $excel = New-Object -ComObject Excel.Application
} catch {
  throw "Could not start Excel COM: $($_.Exception.Message)"
}
try {
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $srcExt = [System.IO.Path]::GetExtension($picked.FullName)
  if ([string]::IsNullOrWhiteSpace($srcExt)) { $srcExt = '.xlsx' }
  $tempCopy = Join-Path $env:TEMP ("bom_check_" + [guid]::NewGuid().ToString('N') + $srcExt)
  Copy-Item -LiteralPath $picked.FullName -Destination $tempCopy -Force
  $wb = $excel.Workbooks.Open($tempCopy, 0, $true)

  $pj = Get-PjmSheetPartDescriptionRows -Workbook $wb -SheetSpec $PjmSheet.Trim()
  Write-Host ("  PJM sheet name: {0}  rows scanned: {1}  Name/Comment column: {2}" -f $pj.SheetName, $pj.Rows.Count, $(if ($pj.HasNameColumn) { 'yes' } else { 'not found (skipping Name vs description check)' }))
  if (-not $pj.HasNameColumn) {
    Write-Warning "PJM sheet '$($pj.SheetName)' has no 'Name', 'Comment', or PJM-prefixed Name/Comment header; skipping Name vs 03-Description General consistency check."
  }

  $dupPnGroups = @($pj.Rows | Where-Object { $_.PartNumber } | Group-Object { $_.PartNumber.Trim() } | Where-Object { $_.Count -gt 1 })
  if ($dupPnGroups.Count -gt 0) {
    Write-Host ''
    Write-Host '=== Duplicate values in 01-PJM Part_Number (rows listed; cells highlighted yellow) ===' -ForegroundColor DarkYellow
    foreach ($g in $dupPnGroups) {
      $rowNums = @($g.Group | ForEach-Object { $_.Row })
      Write-Host ("  '{0}'  Excel rows: {1}" -f $g.Name, ($rowNums -join ', ')) -ForegroundColor DarkYellow
    }
    if ($pj.Worksheet -and $pj.ColPartNumber -gt 0) {
      foreach ($g in $dupPnGroups) {
        foreach ($item in $g.Group) {
          try {
            $pj.Worksheet.Cells.Item($item.Row, $pj.ColPartNumber).Interior.Color = 65535
          } catch {}
        }
      }
      try {
        $wb.Save()
        Copy-Item -LiteralPath $tempCopy -Destination $picked.FullName -Force
        Write-Host "  Workbook saved with duplicate Part_Number cells highlighted." -ForegroundColor DarkYellow
      } catch {
        Write-Warning "Could not save duplicate highlights back to the BOM workbook: $($_.Exception.Message)"
      }
    }
  }

  foreach ($pr in $pj.Rows) {
    $pjn = $pr.PartNumber
    if (-not $pjn) { continue }
    $ctx = Get-ArenaBomExportLineContextCached -PartNumber $pjn
    $linePfx = Get-BomExcelCheckPartLinePrefix -PartNumber $pjn -ExcelRow $pr.Row -DesignatorsCellText $pr.Designators
    if ($ctx.ApiError) {
      $failures.Add(
        "$linePfx Arena API error while looking up item.`n    Altium: $($pr.Description)"
      )
      continue
    }
    if ($ctx.NoItem) {
      $failures.Add(
        "$linePfx Arena has no item with this part number.`n    Altium: $($pr.Description)"
      )
      continue
    }
    $arenaSide = $ctx.SideText
    $arenaMpn = $ctx.Mpn

    $nameMatchesArenaIdentity = $pj.HasNameColumn -and (
      ((-not [string]::IsNullOrWhiteSpace($arenaMpn)) -and (Test-BomNameMatchesManufacturerPartNumber -NameText $pr.Name -MpnText $arenaMpn)) -or
      ((-not [string]::IsNullOrWhiteSpace($ctx.ItemName)) -and (Test-BomNameMatchesManufacturerPartNumber -NameText $pr.Name -MpnText $ctx.ItemName))
    )

    if (-not $nameMatchesArenaIdentity) {
      $chk = Test-BomDescriptionsCompatible -ArenaDescription $arenaSide -AltiumDescription $pr.Description -PartNumber $pjn
      if (-not $chk.Ok) {
        $msg = "$linePfx $($chk.Message)`n    Arena (item desc/name): $arenaSide`n    Altium: $($pr.Description)"
        $failures.Add($msg)
      }
    }

    if ($pj.HasNameColumn -and -not $nameMatchesArenaIdentity -and (Test-ShouldRunNameDescriptionGeneralCheck $pjn)) {
      $nmChk = Test-BomExcelNameVersusDescriptionGeneralCompatible `
        -NameText $pr.Name `
        -DescriptionGeneralText $pr.Description `
        -PartNumber $pjn
      if (-not $nmChk.Ok) {
        $failures.Add(
          "$linePfx $($nmChk.Message)`n    Name: $($pr.Name)`n    03-Description General: $($pr.Description)"
        )
      }
    }
  }
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
  if ($tempCopy -and (Test-Path -LiteralPath $tempCopy)) {
    Remove-Item -LiteralPath $tempCopy -Force -ErrorAction SilentlyContinue
  }
}

if ($failures.Count -gt 0) {
  Write-Host ''
  Write-Host "=== BOM export checks FAILED ($($failures.Count) issue(s)) ===" -ForegroundColor DarkRed
  foreach ($f in $failures) { Write-Host $f; Write-Host '' }
  throw 'BOM export checks failed; not running Convert BOM (Excel -> CSV) v2.'
}

Write-Host ''
Write-Host '=== BOM export checks (Arena + Name vs description): OK ===' -ForegroundColor DarkGreen

if ($CreateCsv) {
  Write-Host ''
  Write-Host '=== Running Convert BOM (Excel -> CSV) v2 ==='
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $convertScript -PartNumber $PartNumber.Trim()
  if ($LASTEXITCODE -ne 0) {
    throw "Convert-AltiumBomAuto.ps1 failed (exit $LASTEXITCODE)."
  }
  Write-Host ''
  Write-Host 'Done. BOM export checked and CSV conversion completed.' -ForegroundColor DarkGreen
}
else {
  Write-Host ''
  Write-Host 'Done. BOM export checks passed (Create CSV was not selected).' -ForegroundColor DarkGreen
}
exit 0
