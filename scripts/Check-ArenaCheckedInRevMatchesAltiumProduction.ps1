<#
.SYNOPSIS
  Verify the PCBA Arena checked-in revision letter matches the Altium production folder.

.DESCRIPTION
  For a PCBA part number:
    1. Reads the most recent checked-in (effective) revision number on the PCBA item.
    2. Resolves the Altium production document folder (Open-AltiumPartFolder -Open ResolveJson).
    3. Compares the revision letter from each (e.g. Arena A1 vs Altium A1 -> letters A and A).

  Read-only. Exit 0 when letters match; exit 2 when they differ.

.PARAMETER PartNumber
  PCBA part number (e.g. 251-10-001 or 251-15-001).

.EXAMPLE
  .\Check-ArenaCheckedInRevMatchesAltiumProduction.ps1 -PartNumber 251-10-001
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $scriptsDir 'AltiumRevisionHelpers.ps1')

$openScript = Join-Path $scriptsDir 'Open-AltiumPartFolder.ps1'
if (-not (Test-Path -LiteralPath $openScript)) { throw "Required script missing: $openScript" }

if ([string]::IsNullOrWhiteSpace($env:ARENA_EMAIL)) { throw 'Set ARENA_EMAIL.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_PASSWORD)) { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_WORKSPACE_ID)) {
  throw 'Set ARENA_WORKSPACE_ID (launcher workspace dropdown sets this per run).'
}

$parts = ConvertTo-PcbaAndBareBoardPartNumbers -PartNumber $PartNumber
$pcbaPn = $parts.Pcba

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

function Get-ResolveJsonFromOpenScript {
  param([string]$Pn)
  $resolveOut = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
    -File $openScript -PartNumber $Pn -Open ResolveJson 2>&1
  if ($LASTEXITCODE -ne 0) {
    throw "Open-AltiumPartFolder.ps1 -Open ResolveJson failed (exit $LASTEXITCODE):`n$($resolveOut -join [Environment]::NewLine)"
  }
  $joined = ($resolveOut | ForEach-Object { "$_" }) -join "`n"
  $rxJson = [regex]'---BEGIN-RESOLVE-JSON---\s*(?<j>\{.*\})\s*---END-RESOLVE-JSON---'
  $jm = $rxJson.Match($joined)
  if (-not $jm.Success) {
    throw "Could not find resolve JSON markers. Output:`n$joined"
  }
  return ($jm.Groups['j'].Value | ConvertFrom-Json)
}

function Get-ArenaCheckedInRevisionNumber {
  param(
    [string]$Pn,
    [hashtable]$Headers,
    [string]$ApiRoot
  )
  $pnEnc = [Uri]::EscapeDataString($Pn)
  $lookup = Invoke-ArenaJson -Uri "$ApiRoot/items?number=$pnEnc&limit=1" -Headers $Headers
  if (-not $lookup.results -or $lookup.results.Count -eq 0) {
    throw "No Arena item found for number '$Pn'."
  }
  $revs = Invoke-ArenaJson -Uri "$ApiRoot/items/$($lookup.results[0].guid)/revisions" -Headers $Headers
  $checkedIn = @($revs.results | Where-Object { $_.status -eq 1 }) |
    Sort-Object { $_.effectiveDateTime } -Descending |
    Select-Object -First 1
  if (-not $checkedIn) {
    throw "No checked-in (effective) revision found for '$Pn'."
  }
  return "$($checkedIn.number)".Trim()
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

Write-Output '=== Arena checked-in revision vs Altium production folder ==='
Write-Output ("  PCBA part number: {0}" -f $pcbaPn)
Write-Output ''

$arenaRevNumber = Get-ArenaCheckedInRevisionNumber -Pn $pcbaPn -Headers $headers -ApiRoot $apiRoot
if ([string]::IsNullOrWhiteSpace($arenaRevNumber)) {
  Write-Output '  ERROR: Arena checked-in revision number is blank.'
  exit 2
}

$arenaLetter = Get-RevisionLetterOnly -RevisionText $arenaRevNumber
Write-Output ("  Arena checked-in rev number: {0}" -f $arenaRevNumber)
Write-Output ("  Arena revision letter:       {0}" -f $arenaLetter)
Write-Output ''

$resolved = Get-ResolveJsonFromOpenScript -Pn $pcbaPn
$prodDocDir = "$($resolved.prodDocDir)".Trim()
$targetDir = "$($resolved.targetDir)".Trim()
$altiumRev = "$($resolved.revLetter)".Trim()

if ([string]::IsNullOrWhiteSpace($prodDocDir)) {
  Write-Output '  ERROR: No Altium production document folder found under any revision.'
  exit 2
}

Write-Output ("  Altium rev folder:           {0}" -f $(if ($targetDir) { $targetDir } else { '(unknown)' }))
Write-Output ("  Altium production folder:    {0}" -f $prodDocDir)

if ([string]::IsNullOrWhiteSpace($altiumRev)) {
  Write-Output '  ERROR: Could not determine revision from the Altium production folder.'
  exit 2
}

$altiumLetter = Get-RevisionLetterOnly -RevisionText $altiumRev
Write-Output ("  Altium revision (folder):    {0}" -f $altiumRev)
Write-Output ("  Altium revision letter:      {0}" -f $altiumLetter)
Write-Output ''

if ($arenaLetter -ceq $altiumLetter) {
  Write-Output ("  -> Revision letters match ({0})." -f $arenaLetter)
  exit 0
}

Write-Output ("  WARN: revision letter mismatch (Arena={0}, Altium={1})." -f $arenaLetter, $altiumLetter)
exit 2
