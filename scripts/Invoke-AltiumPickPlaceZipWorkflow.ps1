<#
.SYNOPSIS
  Optional Pick & Place zip build and/or Arena upload (launcher combined card).

.PARAMETER PartNumber
  Altium part lookup (PPP-XX-XXX).

.PARAMETER CreateZip
  Run Build Pick & Place zip (Open-AltiumPartFolder -Open PickPlaceZip).

.PARAMETER UploadToArena
  Upload or replace "Pick and Place" on the PCBA Arena working revision.

.PARAMETER UseExistingFile
  When uploading: use the zip already in the production folder; do not build if missing.

.PARAMETER RevisionOverride
  Optional revision for the zip name (letter + number, e.g. D1). The letter must match the
  revision letter from Altium outputs; the number you specify is used in the zip file name.
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [string]$RevisionOverride = '',
  [switch]$CreateZip,
  [switch]$UploadToArena,
  [switch]$UseExistingFile,
  [switch]$ForceUpload,
  [switch]$ForceZipOverwrite
)

$ErrorActionPreference = 'Stop'
$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $scriptsDir 'LauncherNestedScriptHelpers.ps1')
$openScript = Join-Path $scriptsDir 'Open-AltiumPartFolder.ps1'
$syncScript = Join-Path $scriptsDir 'Sync-ArenaPickPlaceZip.ps1'
foreach ($p in @($openScript, $syncScript)) {
  if (-not (Test-Path -LiteralPath $p)) { throw "Required script missing: $p" }
}

if (-not $CreateZip -and -not $UploadToArena) {
  throw 'Enable at least one option: Create Zip and/or Upload Zip File to Arena.'
}
if ($UseExistingFile -and -not $UploadToArena) {
  Write-Warning 'Use Existing File only applies when Upload Zip File to Arena is enabled; ignoring.'
  $UseExistingFile = $false
}

$pn = $PartNumber.Trim()

if ($CreateZip) {
  Write-Host '=== Create Zip (Build Pick & Place zip) ==='
  $openArgs = @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass',
    '-File', $openScript,
    '-PartNumber', $pn,
    '-Open', 'PickPlaceZip'
  )
  if (-not [string]::IsNullOrWhiteSpace($RevisionOverride)) {
    $openArgs += '-RevisionOverride', $RevisionOverride.Trim()
  }
  if ($ForceZipOverwrite) { $openArgs += '-ForceZipOverwrite' }
  Invoke-NestedLauncherPowerShell -ArgumentList $openArgs -StepLabel 'Build Pick & Place zip' | Out-Null
}

if ($UploadToArena) {
  Write-Host ''
  Write-Host '=== Upload Zip File to Arena (Pick & Place) ==='
  $syncArgs = @('-File', $syncScript, '-PartNumber', $pn)
  if (-not [string]::IsNullOrWhiteSpace($RevisionOverride)) {
    $syncArgs += '-RevisionOverride', $RevisionOverride.Trim()
  }
  if ($CreateZip -or $UseExistingFile) {
    $syncArgs += '-SkipZipBuildIfMissing'
  }
  if ($ForceUpload) {
    $syncArgs += '-ForceUpload'
  }
  $syncInvokeArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass') + $syncArgs
  Invoke-NestedLauncherPowerShell -ArgumentList $syncInvokeArgs -StepLabel 'Upload Pick & Place zip to Arena' | Out-Null
}

Write-Host ''
Write-Host 'Done.' -ForegroundColor Green
exit 0
