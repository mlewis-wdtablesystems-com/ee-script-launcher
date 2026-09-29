<#
.SYNOPSIS
  Export an Arena assembly BOM to Excel using the SUPPLIER format.

.DESCRIPTION
  Logs into Arena, resolves the item by part number, selects a revision (effective
  latest by default), and requests the native BOM export (format SUPPLIER, Excel).
  Writes the .xlsx file under -OutputDirectory (default: your Documents folder).

  Credentials: ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID.

.PARAMETER PartNumber
  Assembly part number (e.g. 294-10-001).

.PARAMETER RevisionMode
  effective_latest = newest effective (released) revision (default).
  working          = working revision (status 0).

.PARAMETER BomLevels
  SINGLE or FULL (default FULL — fully exploded BOM).

.PARAMETER OutputDirectory
  Folder for the output .xlsx (created if missing).

.EXAMPLE
  .\Export-ArenaSupplierBom.ps1 -PartNumber 294-10-001
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [ValidateSet('working', 'effective_latest')] [string]$RevisionMode = 'effective_latest',
  [ValidateSet('SINGLE', 'FULL')] [string]$BomLevels = 'FULL',
  [string]$OutputDirectory = '',
  [string]$Email = $env:ARENA_EMAIL,
  [string]$WorkspaceId = $env:ARENA_WORKSPACE_ID,
  [string]$BaseUrl = 'https://api.arenasolutions.com/v1'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

if ([string]::IsNullOrWhiteSpace($Email)) { throw 'Set ARENA_EMAIL.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_PASSWORD)) { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($WorkspaceId)) { throw 'Set ARENA_WORKSPACE_ID.' }

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
  $OutputDirectory = [Environment]::GetFolderPath('MyDocuments')
}
$OutputDirectory = $OutputDirectory.Trim()
if (-not (Test-Path -LiteralPath $OutputDirectory)) {
  New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
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

function Invoke-ArenaRaw {
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
  return Invoke-WebRequest @params
}

function Test-BytesLookLikeExcelOrZip([byte[]]$Bytes) {
  if (-not $Bytes -or $Bytes.Length -lt 4) { return $false }
  # XLSX and ZIP both start with PK
  return ($Bytes[0] -eq 0x50 -and $Bytes[1] -eq 0x4B)
}

function Get-ResponseBytes([object]$Resp) {
  if ($Resp.RawContentStream) {
    return $Resp.RawContentStream.ToArray()
  }
  if ($Resp.Content -is [byte[]]) { return $Resp.Content }
  if ($null -ne $Resp.Content) {
    return [System.Text.Encoding]::UTF8.GetBytes([string]$Resp.Content)
  }
  return @()
}

function Get-SuggestedFileName {
  param([string]$PartNumber, [string]$RevNumber, [string]$ContentDisposition)
  if ($ContentDisposition -and ($ContentDisposition -match 'filename=')) {
    $tail = ($ContentDisposition -split 'filename=', 2)[1]
    if ($tail) {
      $fn = ($tail -split ';')[0].Trim().Trim('"')
      if ($fn) { return $fn }
    }
  }
  $safePn = $PartNumber.Trim()
  foreach ($c in [System.IO.Path]::GetInvalidFileNameChars()) {
    $safePn = $safePn.Replace([string]$c, '_')
  }
  $safeRev = if ($RevNumber) { "_Rev$($RevNumber.Trim())" } else { '' }
  return "${safePn}_Supplier_BOM${safeRev}.xlsx"
}

function Select-ItemRevision {
  param([object[]]$RevResults, [string]$ModeKey)
  if ($ModeKey -eq 'working') {
    $w = $RevResults | Where-Object { $_.status -eq 0 } | Select-Object -First 1
    if (-not $w) { throw 'No working revision (status 0) on this item.' }
    return $w
  }
  $eff = @($RevResults | Where-Object { $_.status -eq 1 })
  if ($eff.Count -eq 0) { throw 'No effective (released) revision (status 1). Use -RevisionMode working or release the item first.' }
  return $eff | Sort-Object { $_.effectiveDateTime } -Descending | Select-Object -First 1
}

function Try-DownloadBomExportFromArena {
  param(
    [string]$ApiRoot,
    [hashtable]$Headers,
    [string]$ItemGuid,
    [string]$RevisionGuid,
    [string]$BomLevelsKey
  )
  $uris = @(
    "$ApiRoot/items/$RevisionGuid/bom/export",
    "$ApiRoot/items/$ItemGuid/bom/export"
  )
  $bodies = @(
    @{ format = 'SUPPLIER'; fileFormat = 'EXCEL'; bomLevels = $BomLevelsKey } | ConvertTo-Json -Compress
    @{ bomFormat = 'SUPPLIER'; fileFormat = 'EXCEL'; bomLevels = $BomLevelsKey } | ConvertTo-Json -Compress
    @{ format = 'SUPPLIER'; fileType = 'EXCEL'; bomLevels = $BomLevelsKey } | ConvertTo-Json -Compress
    @{ bomExportFormat = 'SUPPLIER'; exportFileType = 'EXCEL'; bomLevels = $BomLevelsKey } | ConvertTo-Json -Compress
    @{ format = 'SUPPLIER'; fileFormat = 'XLSX'; bomLevels = $BomLevelsKey } | ConvertTo-Json -Compress
  )

  $attempts = New-Object System.Collections.Generic.List[string]

  foreach ($uri in $uris) {
    $q = "$uri`?format=SUPPLIER&fileFormat=EXCEL&bomLevels=$BomLevelsKey"
    $attempts.Add("GET $q")
    try {
      $r = Invoke-ArenaRaw -Uri $q -Method Get -Headers $Headers
      $bytes = Get-ResponseBytes $r
      $ct = "$($r.Headers['Content-Type'])"
      if (Test-BytesLookLikeExcelOrZip $bytes) {
        return @{
          Bytes = $bytes
          ContentType = $ct
          ContentDisposition = "$($r.Headers['Content-Disposition'])"
          Endpoint = $q
        }
      }
    }
    catch {
      $attempts.Add("  GET failed: $($_.Exception.Message)")
    }

    foreach ($body in $bodies) {
      $attempts.Add("POST $uri body=$body")
      try {
        $r = Invoke-ArenaRaw -Uri $uri -Method Post -Headers $Headers -Body $body
        $bytes = Get-ResponseBytes $r
        $ct = "$($r.Headers['Content-Type'])"
        if (Test-BytesLookLikeExcelOrZip $bytes) {
          return @{
            Bytes = $bytes
            ContentType = $ct
            ContentDisposition = "$($r.Headers['Content-Disposition'])"
            Endpoint = "$uri POST $body"
          }
        }
        $text = [System.Text.Encoding]::UTF8.GetString($bytes)
        if ($text -and $text.TrimStart().StartsWith('{')) {
          $j = $text | ConvertFrom-Json
          if ($j.errors) {
            $attempts.Add("  Arena: " + (($j.errors | ForEach-Object { $_.message }) -join '; '))
          }
        }
      }
      catch {
        $msg = $_.Exception.Message
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
          $msg = "$msg`n$($_.ErrorDetails.Message)"
        }
        $attempts.Add("  POST failed: $msg")
      }
    }
  }

  return @{ Attempts = $attempts }
}

# ---------- Login ----------
$apiRoot = $BaseUrl.TrimEnd('/')
$pn = $PartNumber.Trim()

$loginBody = @{
  email       = $Email
  password    = $env:ARENA_PASSWORD
  workspaceId = [long]$WorkspaceId
} | ConvertTo-Json

$login = Invoke-ArenaJson -Uri "$apiRoot/login" -Method Post -Body $loginBody
$sessionId = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sessionId)) { $sessionId = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sessionId)) { throw 'Login response did not include arenaSessionId.' }

$headers = @{
  arena_session_id = $sessionId
  'Content-Type'   = 'application/json'
}

# ---------- Resolve item ----------
$pnEnc = [Uri]::EscapeDataString($pn)
$lookup = Invoke-ArenaJson -Uri "$apiRoot/items?number=$pnEnc&limit=1" -Headers $headers
if (-not $lookup.results -or $lookup.results.Count -eq 0) {
  throw "No Arena item found for part number '$pn'."
}
$item = $lookup.results[0]
$itemGuid = "$($item.guid)".Trim()
if (-not $itemGuid) { throw 'Item lookup returned no guid.' }

$revs = Invoke-ArenaJson -Uri "$apiRoot/items/$itemGuid/revisions" -Headers $headers
if (-not $revs.results -or $revs.results.Count -eq 0) {
  throw "No revisions returned for item '$pn'."
}
$rev = Select-ItemRevision -RevResults @($revs.results) -ModeKey $RevisionMode
$revGuid = "$($rev.guid)".Trim()
$revNum = if ($null -ne $rev.number) { "$($rev.number)".Trim() } else { '' }

Write-Host "=== Export Supplier BOM ==="
Write-Host ("  Part:     {0}" -f $pn)
Write-Host ("  Item:     {0}" -f $itemGuid)
Write-Host ("  Revision: {0} (guid {1}, mode {2})" -f $revNum, $revGuid, $RevisionMode)
Write-Host ("  Format:   SUPPLIER / Excel, bomLevels={0}" -f $BomLevels)

# ---------- Export ----------
$result = Try-DownloadBomExportFromArena -ApiRoot $apiRoot -Headers $headers `
  -ItemGuid $itemGuid -RevisionGuid $revGuid -BomLevelsKey $BomLevels

if (-not $result.Bytes) {
  $log = if ($result.Attempts) { ($result.Attempts | ForEach-Object { "  $_" }) -join "`n" } else { '  (none)' }
  throw @"
Arena did not return an Excel BOM export. Tried native BOM export endpoints with format SUPPLIER.

$log

If your workspace uses a different API shape, export once from the Arena UI (BOM view, Export, Supplier, Excel)
and capture the network request, or contact Arena support for the BOM export endpoint.
"@
}

$fileName = Get-SuggestedFileName -PartNumber $pn -RevNumber $revNum -ContentDisposition $result.ContentDisposition
$outPath = Join-Path $OutputDirectory $fileName
[System.IO.File]::WriteAllBytes($outPath, $result.Bytes)

Write-Host ''
Write-Host ("Saved: {0}" -f $outPath) -ForegroundColor DarkGreen
Write-Host ("  Size:  {0:N0} bytes" -f $result.Bytes.Length)
Write-Host ("  Via:   {0}" -f $result.Endpoint)
exit 0
