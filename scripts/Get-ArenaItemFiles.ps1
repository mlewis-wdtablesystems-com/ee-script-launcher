<#
.SYNOPSIS
  Print, as JSON, the files attached to a given item revision.

.DESCRIPTION
  Resolves the working (default) or effective revision of -PartNumber and
  emits a JSON document the launcher's discovery flow can consume:

    { "rev": { "guid": "...", "status": 0, "number": null },
      "files": [
        { "association_guid": "...",
          "file_guid": "...",
          "number": "FILE-001234",
          "title": "...",
          "edition": "3",
          "format": "pdf",
          "category_path": "File\\Internal File\\..." }
      ] }

  Credentials via env: ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID.
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [ValidateSet('working','effective')] [string]$Revision = 'working',
  [string]$Email = $env:ARENA_EMAIL,
  [string]$WorkspaceId = $env:ARENA_WORKSPACE_ID,
  [string]$BaseUrl = 'https://api.arenasolutions.com/v1'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

function Invoke-ArenaJson {
  param([string]$Uri, [string]$Method = 'Get', [hashtable]$Headers, [string]$Body, [string]$ContentType = 'application/json')
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
if ([string]::IsNullOrWhiteSpace($password))    { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($WorkspaceId)) { throw 'Set ARENA_WORKSPACE_ID or pass -WorkspaceId.' }

$apiRoot = $BaseUrl.TrimEnd('/')

$loginBody = @{ email = $Email; password = $password; workspaceId = [long]$WorkspaceId } | ConvertTo-Json
try {
  $login = Invoke-ArenaJson -Uri "$apiRoot/login" -Method Post -Body $loginBody
} catch { throw "Arena login failed: $($_.Exception.Message)" }

$sid = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sid)) { $sid = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sid)) { throw 'Login response did not include arenaSessionId.' }

$headers = @{ arena_session_id = $sid; 'Content-Type' = 'application/json' }

$pnEnc = [Uri]::EscapeDataString($PartNumber.Trim())
$lookup = Invoke-ArenaJson -Uri "$apiRoot/items?number=$pnEnc&limit=1" -Headers $headers
if (-not $lookup.results -or $lookup.results.Count -eq 0) {
  throw "No item found for number '$($PartNumber.Trim())'."
}
$anyGuid = $lookup.results[0].guid

$revs = Invoke-ArenaJson -Uri "$apiRoot/items/$anyGuid/revisions" -Headers $headers
$wantedStatus = if ($Revision -eq 'working') { 0 } else { 1 }
$rev = $revs.results | Where-Object { $_.status -eq $wantedStatus } | Select-Object -First 1
if (-not $rev) { throw "No '$Revision' revision found for $($PartNumber.Trim())." }

$resp = Invoke-ArenaJson -Uri "$apiRoot/items/$($rev.guid)/files" -Headers $headers

$files = @()
if ($resp -and $resp.results) {
  foreach ($a in $resp.results) {
    $files += [pscustomobject]@{
      association_guid = $a.guid
      file_guid        = $a.file.guid
      number           = $a.file.number
      title            = $a.file.title
      edition          = $a.file.edition
      format           = $a.file.format
      category_path    = $a.file.category.path
      category_name    = $a.file.category.name
      locked           = $a.file.locked
    }
  }
}

[pscustomobject]@{
  rev = [pscustomobject]@{
    guid   = $rev.guid
    status = $rev.status
    number = $rev.number
  }
  files = $files
} | ConvertTo-Json -Depth 6 -Compress
