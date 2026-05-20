<#
.SYNOPSIS
  Look up an Arena PLM item by internal part number via the REST API.

.DESCRIPTION
  POST /v1/login with email, password, and workspaceId, then GET /v1/items
  with query number=<PartNumber>. Uses header arena_session_id on API calls
  after login; the JSON token field is arenaSessionId (see PTC Arena REST docs).

  Login response may also recommend header Arena-Usage-Reason for integration
  identification; add that later if your workspace requires it.

  Credentials: set environment variables (recommended):
    ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID

  You may pass -Email and -WorkspaceId explicitly; password is always read
  from ARENA_PASSWORD so it is not committed to scripts or shell history via
  parameters.

.PARAMETER PartNumber
  Internal item number to search (e.g. 294-10-001).

.PARAMETER ResponseView
  Arena items responseview: minimum, compact, or full (default full).

.PARAMETER BaseUrl
  API base URL (default https://api.arenasolutions.com/v1).

.EXAMPLE
  $env:ARENA_EMAIL = 'you@company.com'
  $env:ARENA_PASSWORD = '***'
  $env:ARENA_WORKSPACE_ID = '123456789'
  .\Get-ArenaItem.ps1 -PartNumber '294-10-001'
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [ValidateSet('minimum', 'compact', 'full')] [string]$ResponseView = 'full',
  [ValidateSet('text', 'json', 'object')] [string]$Format = 'text',
  [string]$Email = $env:ARENA_EMAIL,
  [string]$WorkspaceId = $env:ARENA_WORKSPACE_ID,
  [string]$BaseUrl = 'https://api.arenasolutions.com/v1'
)

$ErrorActionPreference = 'Stop'

# Emit UTF-8 so callers (terminals, web UIs, pipes) get correct ± etc.
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

# Windows PowerShell 5.1 Invoke-RestMethod misreads UTF-8 JSON as CP1252.
# Use Invoke-WebRequest + explicit UTF-8 decode so descriptions stay intact.
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
if ([string]::IsNullOrWhiteSpace($Email)) {
  throw 'Set ARENA_EMAIL or pass -Email.'
}
if ([string]::IsNullOrWhiteSpace($password)) {
  throw 'Set ARENA_PASSWORD (required for login).'
}
if ([string]::IsNullOrWhiteSpace($WorkspaceId)) {
  throw 'Set ARENA_WORKSPACE_ID or pass -WorkspaceId.'
}

$loginBody = @{
  email       = $Email
  password    = $password
  workspaceId = [long]$WorkspaceId
} | ConvertTo-Json

try {
  $login = Invoke-ArenaJson -Uri "$BaseUrl/login" -Method Post -Body $loginBody
} catch {
  throw "Arena login failed: $($_.Exception.Message)"
}

$sessionId = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sessionId)) {
  $sessionId = $login.arena_session_id
}
if ([string]::IsNullOrWhiteSpace($sessionId)) {
  throw 'Login response did not include arenaSessionId (or arena_session_id).'
}

$headers = @{
  arena_session_id = $sessionId
  'Content-Type'   = 'application/json'
}

$pnEnc = [Uri]::EscapeDataString($PartNumber.Trim())
$viewEnc = [Uri]::EscapeDataString($ResponseView)
$uri = "$BaseUrl/items?number=$pnEnc&limit=50&responseview=$viewEnc"

try {
  $items = Invoke-ArenaJson -Uri $uri -Headers $headers
} catch {
  throw "Arena GET items failed: $($_.Exception.Message)"
}

if (-not $items.results -or $items.results.Count -eq 0) {
  Write-Warning "No items returned for number='$($PartNumber.Trim())'."
  exit 2
}

switch ($Format) {
  'json' {
    [pscustomobject]@{
      PartNumber = $PartNumber.Trim()
      Count      = $items.count
      Results    = $items.results
    } | ConvertTo-Json -Depth 25
  }
  'object' {
    [pscustomobject]@{
      PartNumber = $PartNumber.Trim()
      Count      = $items.count
      Results    = $items.results
    }
  }
  default {
    Write-Output ("Part: {0}    Count: {1}" -f $PartNumber.Trim(), $items.count)
    Write-Output ''
    foreach ($it in $items.results) {
      $lines = @(
        ("Number:         {0}" -f $it.number),
        ("Name:           {0}" -f $it.name),
        ("Revision:       {0}  ({1})" -f $it.revisionNumber, $it.revisionStatus),
        ("Lifecycle:      {0}" -f ($it.lifecyclePhase.name)),
        ("Category:       {0}" -f ($it.category.name)),
        ("Assembly type:  {0}" -f $it.assemblyType),
        ("UOM:            {0}" -f $it.uom),
        ("Owner:          {0}" -f ($it.owner.fullName)),
        ("Effective:      {0}" -f $it.effectiveDateTime),
        ("Created:        {0}" -f $it.creationDateTime),
        ("GUID:           {0}" -f $it.guid),
        ("API URL:        {0}" -f $it.url.api),
        ("App URL:        {0}" -f $it.url.app)
      )
      Write-Output ($lines -join [Environment]::NewLine)
      Write-Output ''
    }
  }
}

exit 0
