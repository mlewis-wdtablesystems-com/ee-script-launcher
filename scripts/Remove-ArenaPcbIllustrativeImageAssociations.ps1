<#
.SYNOPSIS
  Remove Illustrative Image file associations from a bare PCB Arena item.

.DESCRIPTION
  For a PCBA part number, maps to the bare board (PPP-20-XXX / PPP-25-XXX) and
  removes every Illustrative Image file linked on that item's working revision.
  Product images belong on the PCBA item only.

  Uses DELETE /items/<revGuid>/files/<fileGuid>. The file object remains in Arena;
  only the association on the working revision is removed.

.PARAMETER PartNumber
  PCBA part number (e.g. 251-10-001 or 251-15-001).

.EXAMPLE
  .\Remove-ArenaPcbIllustrativeImageAssociations.ps1 -PartNumber 251-10-001
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $scriptsDir 'AltiumRevisionHelpers.ps1')

if ([string]::IsNullOrWhiteSpace($env:ARENA_EMAIL)) { throw 'Set ARENA_EMAIL.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_PASSWORD)) { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_WORKSPACE_ID)) {
  throw 'Set ARENA_WORKSPACE_ID (launcher workspace dropdown sets this per run).'
}

$parts = ConvertTo-PcbaAndBareBoardPartNumbers -PartNumber $PartNumber
$barePn = $parts.Bare

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

function Invoke-ArenaDelete {
  param(
    [Parameter(Mandatory = $true)] [string]$Uri,
    [hashtable]$Headers
  )
  $params = @{ Uri = $Uri; Method = 'Delete'; Headers = $Headers; UseBasicParsing = $true }
  return Invoke-WebRequest @params
}

function Test-IsIllustrativeImageCategory {
  param([object]$Association)
  if (-not $Association -or -not $Association.file -or -not $Association.file.category) { return $false }
  $cat = $Association.file.category
  $n = $cat.name
  if ($null -ne $n -and $n.ToString().Trim().ToLowerInvariant() -eq 'illustrative image') {
    return $true
  }
  $path = $cat.path
  if ($null -ne $path -and $path.ToString().ToLowerInvariant().EndsWith('\illustrative image')) {
    return $true
  }
  return $false
}

function Test-ArenaFileLocked {
  param([object]$Association)
  if (-not $Association -or -not $Association.file) { return $false }
  $v = $Association.file.locked
  if ($null -eq $v) { return $false }
  if ($v -is [bool]) { return $v }
  $s = ("$v").Trim().ToLowerInvariant()
  return $s -in @('true', '1', 'yes')
}

$apiRoot = 'https://api.arenasolutions.com/v1'
$loginBody = @{
  email       = $env:ARENA_EMAIL
  password    = $env:ARENA_PASSWORD
  workspaceId = [long]$env:ARENA_WORKSPACE_ID
} | ConvertTo-Json
try {
  $login = Invoke-ArenaJson -Uri "$apiRoot/login" -Method Post -Body $loginBody
} catch { throw "Arena login failed: $($_.Exception.Message)" }

$sid = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sid)) { $sid = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sid)) { throw 'Login response did not include arenaSessionId.' }

$headers = @{ arena_session_id = $sid; 'Content-Type' = 'application/json' }

$pnEnc = [Uri]::EscapeDataString($barePn)
$lookup = Invoke-ArenaJson -Uri "$apiRoot/items?number=$pnEnc&limit=1" -Headers $headers
if (-not $lookup.results -or $lookup.results.Count -eq 0) {
  throw "No Arena item found for PCB part number '$barePn'."
}

$revs = Invoke-ArenaJson -Uri "$apiRoot/items/$($lookup.results[0].guid)/revisions" -Headers $headers
$working = $revs.results | Where-Object { $_.status -eq 0 } | Select-Object -First 1
if (-not $working) {
  throw "No working revision (status 0) on PCB item $barePn."
}

Write-Output '=== Remove Illustrative Image from PCB (bare board) ==='
Write-Output ("  PCBA part number:     {0}" -f $parts.Pcba)
Write-Output ("  PCB part number:      {0}" -f $barePn)
Write-Output ("  Working revision:     {0}" -f $(if ($working.number) { $working.number } else { '(no number)' }))
Write-Output ''

$filesResp = Invoke-ArenaJson -Uri "$apiRoot/items/$($working.guid)/files" -Headers $headers
$illustrative = @($filesResp.results | Where-Object { Test-IsIllustrativeImageCategory $_ })

if ($illustrative.Count -eq 0) {
  Write-Output '  No Illustrative Image files linked on PCB working revision.'
  exit 0
}

Write-Output ("  Found {0} Illustrative Image file(s) to remove:" -f $illustrative.Count)
$removed = 0
$failed = New-Object System.Collections.Generic.List[string]

foreach ($assoc in $illustrative) {
  $title = $assoc.file.title
  $fileGuid = "$($assoc.file.guid)".Trim()
  $line = "    - {0} ({1})" -f $title, $assoc.file.number
  if (Test-ArenaFileLocked -Association $assoc) {
    Write-Output ($line + '  SKIP: file is locked.')
    [void]$failed.Add("Locked Illustrative Image '$title' on $barePn.")
    continue
  }
  try {
    $resp = Invoke-ArenaDelete -Uri "$apiRoot/items/$($working.guid)/files/$fileGuid" -Headers $headers
    $code = [int]$resp.StatusCode
    if ($code -ne 204 -and $code -ne 200) {
      throw "Unexpected HTTP $code"
    }
    Write-Output ($line + '  -> removed.')
    $removed++
  } catch {
    $msg = $_.Exception.Message
    Write-Output ($line + "  ERROR: $msg")
    [void]$failed.Add("Could not remove Illustrative Image '$title' on ${barePn}: $msg")
  }
}

Write-Output ''
Write-Output ("  Removed {0} of {1} Illustrative Image association(s)." -f $removed, $illustrative.Count)

if ($failed.Count -gt 0) {
  foreach ($p in $failed) { Write-Output ("  WARN: {0}" -f $p) }
  exit 2
}

exit 0
