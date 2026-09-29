<#
.SYNOPSIS
  Set working revision numbers on a PCBA and its mapped bare-board PCB in Arena.

.DESCRIPTION
  For a PCBA part number (PPP-10-XXX / PPP-15-XXX) and revision (letter + number, e.g. A1):

    - PCBA working revision -> full revision (e.g. A1)
    - Bare board (PPP-20/25-XXX) working revision -> letter only (e.g. A)

  Uses Set-ArenaItemWorkingRevisionNumber.ps1 for each item.

.PARAMETER PartNumber
  PCBA part number (e.g. 239-10-001).

.PARAMETER Revision
  Revision with letter and number (e.g. A1 or A 1).

.PARAMETER ConfirmOverwrite
  Skip the overwrite confirmation (Script Launcher sets this after you choose Confirm).

.EXAMPLE
  .\Set-ArenaPcbaWorkingRevisions.ps1 -PartNumber 239-10-001 -Revision A1
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [Parameter(Mandatory = $true)] [string]$Revision,
  [switch]$ConfirmOverwrite,
  [string]$Email = $env:ARENA_EMAIL,
  [string]$WorkspaceId = $env:ARENA_WORKSPACE_ID,
  [string]$BaseUrl = 'https://api.arenasolutions.com/v1'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $scriptsDir 'AltiumRevisionHelpers.ps1')

$setRevScript = Join-Path $scriptsDir 'Set-ArenaItemWorkingRevisionNumber.ps1'
if (-not (Test-Path -LiteralPath $setRevScript)) {
  throw "Required script missing: $setRevScript"
}

if ([string]::IsNullOrWhiteSpace($env:ARENA_EMAIL)) { throw 'Set ARENA_EMAIL.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_PASSWORD)) { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_WORKSPACE_ID)) {
  throw 'Set ARENA_WORKSPACE_ID (launcher workspace dropdown sets this per run).'
}

$password = $env:ARENA_PASSWORD

$revParsed = Parse-AltiumRevisionToken -RevisionText $Revision
if (-not $revParsed.Number) {
  throw "Revision must include a number after the letter (e.g. A1). Got: '$Revision'."
}

$parts = ConvertTo-PcbaAndBareBoardPartNumbers -PartNumber $PartNumber
$pcbaPn = $parts.Pcba
$barePn = $parts.Bare
$pcbaRev = $revParsed.Display
$bareRev = $revParsed.Letter

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

function Get-WorkingRevisionNumber {
  param(
    [Parameter(Mandatory = $true)] [string]$ItemNumber,
    [Parameter(Mandatory = $true)] [hashtable]$Headers,
    [Parameter(Mandatory = $true)] [string]$ApiRoot
  )
  $pnEnc = [Uri]::EscapeDataString($ItemNumber.Trim())
  $lookup = Invoke-ArenaJson -Uri "$ApiRoot/items?number=$pnEnc&limit=1" -Headers $Headers
  if (-not $lookup.results -or $lookup.results.Count -eq 0) {
    throw "No Arena item found for number '$ItemNumber'."
  }
  $revs = Invoke-ArenaJson -Uri "$ApiRoot/items/$($lookup.results[0].guid)/revisions" -Headers $Headers
  $working = $revs.results | Where-Object { $_.status -eq 0 } | Select-Object -First 1
  if (-not $working) { throw "No working revision (status 0) for item '$ItemNumber'." }
  if ($null -eq $working.number) { return '' }
  return ("$($working.number)").Trim()
}

function Test-NeedsOverwriteConfirm {
  param(
    [string]$Current,
    [Parameter(Mandatory = $true)] [string]$Target
  )
  $cur = if ($null -eq $Current) { '' } else { ("$Current").Trim() }
  if ([string]::IsNullOrWhiteSpace($cur)) { return $false }
  return ($cur -cne $Target)
}

function Format-RevisionForPrompt([string]$Value) {
  if ([string]::IsNullOrWhiteSpace($Value)) { return '(blank)' }
  return $Value.Trim()
}

$apiRoot = $BaseUrl.TrimEnd('/')
$loginBody = @{ email = $Email; password = $password; workspaceId = [long]$WorkspaceId } | ConvertTo-Json
$login = Invoke-ArenaJson -Uri "$apiRoot/login" -Method Post -Body $loginBody
$sid = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sid)) { $sid = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sid)) { throw 'Login response did not include arenaSessionId.' }
$headers = @{ arena_session_id = $sid; 'Content-Type' = 'application/json' }

$pcbaCurrent = Get-WorkingRevisionNumber -ItemNumber $pcbaPn -Headers $headers -ApiRoot $apiRoot
$bareCurrent = Get-WorkingRevisionNumber -ItemNumber $barePn -Headers $headers -ApiRoot $apiRoot

$needConfirm = (
  (Test-NeedsOverwriteConfirm -Current $pcbaCurrent -Target $pcbaRev) -or
  (Test-NeedsOverwriteConfirm -Current $bareCurrent -Target $bareRev)
)

Write-Host '=== Set PCBA and PCB working revisions ===' -ForegroundColor Cyan
Write-Host ("  PCBA part:              {0}" -f $pcbaPn)
Write-Host ("  Bare board (Arena):     {0}" -f $barePn)
Write-Host ("  Revision (PCBA):        {0}" -f $pcbaRev)
Write-Host ("  Revision (bare letter): {0}" -f $bareRev)
Write-Host ("  Current PCBA working:   {0}" -f $(Format-RevisionForPrompt $pcbaCurrent))
Write-Host ("  Current PCB working:    {0}" -f $(Format-RevisionForPrompt $bareCurrent))
Write-Host ''

if ($needConfirm -and -not $ConfirmOverwrite) {
  $msg = (
    "Current revision is {0} for PCBA and {1} for PCB, are you sure you want to overwrite?" -f
    (Format-RevisionForPrompt $pcbaCurrent),
    (Format-RevisionForPrompt $bareCurrent)
  )
  Write-Host $msg -ForegroundColor DarkYellow
  Write-Host '---UPLOAD-CONFIRM-PROMPT---'
  Write-Host (@{ message = $msg } | ConvertTo-Json -Compress)
  exit 12
}

function Invoke-Step {
  param(
    [Parameter(Mandatory = $true)] [string]$Title,
    [Parameter(Mandatory = $true)] [scriptblock]$Action
  )
  Write-Host "=== $Title ===" -ForegroundColor Cyan
  & $Action
  if ($LASTEXITCODE -ne 0) {
    throw "Step failed (exit $LASTEXITCODE): $Title"
  }
  Write-Host ''
}

Invoke-Step -Title "Set PCBA working revision ($pcbaPn -> $pcbaRev)" -Action {
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $setRevScript `
    -PartNumber $pcbaPn -RevisionNumber $pcbaRev
}

Invoke-Step -Title "Set bare-board working revision ($barePn -> $bareRev)" -Action {
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $setRevScript `
    -PartNumber $barePn -RevisionNumber $bareRev
}

Write-Host '=== Done ===' -ForegroundColor DarkGreen
exit 0
