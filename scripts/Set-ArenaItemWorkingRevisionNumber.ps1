<#
.SYNOPSIS
  Set the working revision number on an Arena item.

.DESCRIPTION
  PUT /v1/items/<workingRevItemGuid> with { revisionNumber: "<value>" }.
  Only the working revision (status 0) is updated.

.PARAMETER PartNumber
  Arena item number (e.g. 239-10-001 or 239-20-001).

.PARAMETER RevisionNumber
  New revision number string (e.g. A1 or A).

.EXAMPLE
  .\Set-ArenaItemWorkingRevisionNumber.ps1 -PartNumber 239-10-001 -RevisionNumber A1
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [Parameter(Mandatory = $true)] [string]$RevisionNumber,
  [string]$Email = $env:ARENA_EMAIL,
  [string]$WorkspaceId = $env:ARENA_WORKSPACE_ID,
  [string]$BaseUrl = 'https://api.arenasolutions.com/v1'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

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
  if ([int]$resp.StatusCode -ge 400) {
    $errBody = ''
    try {
      $bytes = $resp.RawContentStream.ToArray()
      $errBody = [System.Text.Encoding]::UTF8.GetString($bytes)
    } catch {}
    throw ("Arena $Method $Uri returned $($resp.StatusCode): $errBody")
  }
  $bytes = $resp.RawContentStream.ToArray()
  $text = [System.Text.Encoding]::UTF8.GetString($bytes)
  if ([string]::IsNullOrWhiteSpace($text)) { return $null }
  return $text | ConvertFrom-Json
}

$password = $env:ARENA_PASSWORD
if ([string]::IsNullOrWhiteSpace($Email)) { throw 'Set ARENA_EMAIL.' }
if ([string]::IsNullOrWhiteSpace($password)) { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($WorkspaceId)) { throw 'Set ARENA_WORKSPACE_ID.' }

$newRev = ("$RevisionNumber").Trim()
if (-not $newRev) { throw 'RevisionNumber is empty.' }

$apiRoot = $BaseUrl.TrimEnd('/')
$loginBody = @{ email = $Email; password = $password; workspaceId = [long]$WorkspaceId } | ConvertTo-Json
$login = Invoke-ArenaJson -Uri "$apiRoot/login" -Method Post -Body $loginBody
$sid = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sid)) { $sid = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sid)) { throw 'Login response did not include arenaSessionId.' }
$headers = @{ arena_session_id = $sid; 'Content-Type' = 'application/json' }

$pn = $PartNumber.Trim()
$pnEnc = [Uri]::EscapeDataString($pn)
$lookup = Invoke-ArenaJson -Uri "$apiRoot/items?number=$pnEnc&limit=1" -Headers $headers
if (-not $lookup.results -or $lookup.results.Count -eq 0) {
  throw "No Arena item found for number '$pn'."
}

$itemAnyGuid = $lookup.results[0].guid
$revs = Invoke-ArenaJson -Uri "$apiRoot/items/$itemAnyGuid/revisions" -Headers $headers
$working = $revs.results | Where-Object { $_.status -eq 0 } | Select-Object -First 1
if (-not $working) { throw "No working revision (status 0) for item '$pn'." }

$workingGuid = ("$($working.guid)").Trim()
$oldNum = if ($null -ne $working.number) { ("$($working.number)").Trim() } else { '' }

Write-Host ("Item {0}: working revision guid={1}" -f $pn, $workingGuid)
Write-Host ("  Current revision number: {0}" -f $(if ($oldNum) { $oldNum } else { '(none)' }))
Write-Host ("  New revision number:     {0}" -f $newRev)

if ($oldNum -eq $newRev) {
  Write-Host '  (unchanged — skip PUT)' -ForegroundColor DarkGray
  exit 0
}

$body = @{ revisionNumber = $newRev } | ConvertTo-Json -Compress
$updated = Invoke-ArenaJson -Uri "$apiRoot/items/$workingGuid" -Method Put -Headers $headers -Body $body

$after = if ($updated -and $null -ne $updated.revisionNumber) {
  ("$($updated.revisionNumber)").Trim()
} else {
  $newRev
}
Write-Host ("  Updated. revisionNumber={0}" -f $after) -ForegroundColor DarkGreen
exit 0
