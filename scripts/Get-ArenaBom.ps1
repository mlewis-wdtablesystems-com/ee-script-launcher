<#
.SYNOPSIS
  Print the Arena BOM (immediate child items) for a given internal part number.

.DESCRIPTION
  POST /v1/login with email, password, workspaceId, then GET /v1/items?number=<PN>
  to resolve the GUID, then GET /v1/items/<GUID>/bom and emit a flat table.

  Credentials: ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID env vars
  (password is always read from ARENA_PASSWORD).

.PARAMETER PartNumber
  Internal item number (e.g. 294-10-001).

.PARAMETER Format
  table (default), json, or csv.

.EXAMPLE
  .\Get-ArenaBom.ps1 -PartNumber 294-10-001
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [ValidateSet('table', 'json', 'csv')] [string]$Format = 'table',
  [switch]$IncludeAdditionalAttributes,
  [string]$Email = $env:ARENA_EMAIL,
  [string]$WorkspaceId = $env:ARENA_WORKSPACE_ID,
  [string]$BaseUrl = 'https://api.arenasolutions.com/v1'
)

$ErrorActionPreference = 'Stop'

$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

# Windows PowerShell 5.1 Invoke-RestMethod misreads UTF-8 JSON as CP1252.
# Use Invoke-WebRequest + explicit UTF-8 decode so descriptions like "47UH ±10%" stay intact.
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

$password = $env:ARENA_PASSWORD
if ([string]::IsNullOrWhiteSpace($Email))       { throw 'Set ARENA_EMAIL or pass -Email.' }
if ([string]::IsNullOrWhiteSpace($password))    { throw 'Set ARENA_PASSWORD (required for login).' }
if ([string]::IsNullOrWhiteSpace($WorkspaceId)) { throw 'Set ARENA_WORKSPACE_ID or pass -WorkspaceId.' }

$apiRoot = $BaseUrl.TrimEnd('/')

$loginBody = @{
  email       = $Email
  password    = $password
  workspaceId = [long]$WorkspaceId
} | ConvertTo-Json

try {
  $login = Invoke-ArenaJson -Uri "$apiRoot/login" -Method Post -Body $loginBody
} catch {
  throw "Arena login failed: $($_.Exception.Message)"
}

$sid = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sid)) { $sid = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sid)) { throw 'Login response did not include arenaSessionId.' }

$headers = @{
  arena_session_id = $sid
  'Content-Type'   = 'application/json'
}

$pnEnc = [Uri]::EscapeDataString($PartNumber.Trim())
$lookup = Invoke-ArenaJson -Uri "$apiRoot/items?number=$pnEnc&limit=1" -Headers $headers
if (-not $lookup.results -or $lookup.results.Count -eq 0) {
  throw "No item found for number '$($PartNumber.Trim())'."
}

$parent = $lookup.results[0]
$guid = $parent.guid

$bomUri = "$apiRoot/items/$guid/bom"
if ($IncludeAdditionalAttributes) { $bomUri += '?includeAdditionalAttributes=true' }
$bom = Invoke-ArenaJson -Uri $bomUri -Headers $headers

$rows = foreach ($r in $bom.results) {
  [pscustomobject]@{
    LineNumber  = $r.lineNumber
    Quantity    = $r.quantity
    PartNumber  = $r.item.number
    Revision    = $r.item.revisionNumber
    Name        = $r.item.name
    RefDes      = $r.refDes
    Notes       = $r.notes
    ItemGuid    = $r.item.guid
    AppUrl      = $r.item.url.app
  }
}

switch ($Format) {
  'json' {
    [pscustomobject]@{
      Parent = [pscustomobject]@{
        Number     = $parent.number
        Name       = $parent.name
        Revision   = $parent.revisionNumber
        Guid       = $parent.guid
        Url        = $parent.url.app
      }
      Count = $bom.count
      Lines = $rows
    } | ConvertTo-Json -Depth 10
  }
  'csv' {
    $rows | ConvertTo-Csv -NoTypeInformation
  }
  default {
    Write-Output ("Parent: {0}  {1}  Rev {2}" -f $parent.number, $parent.name, $parent.revisionNumber)
    Write-Output ("BOM lines: {0}" -f $bom.count)
    Write-Output ''
    $rows |
      Select-Object LineNumber, Quantity, PartNumber, Revision, Name, RefDes |
      Format-Table -AutoSize |
      Out-String -Width 240
  }
}

if (-not $rows) { exit 2 }
exit 0
