<#
.SYNOPSIS
  Optional fabrication zip build and/or Arena Gerber zip upload (launcher combined card).

.PARAMETER PartNumber
  Altium part lookup (PPP-XX-XXX).

.PARAMETER CreateZip
  Run Check and Build Fabrication Zip (Open-AltiumPartFolder -Open FabricationZip).

.PARAMETER UploadToArena
  Upload or replace "Altium Gerber Files" on the bare-board Arena working revision.

.PARAMETER UseExistingFile
  When uploading: use the zip already in the production folder; do not build if missing.
  Ignored unless UploadToArena is set. When CreateZip is also set, the zip is built first.
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
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
$syncScript = Join-Path $scriptsDir 'Sync-ArenaFabricationGerbers.ps1'
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
  Write-Host '=== Create Zip (Check and Build Fabrication Zip) ==='
  $openArgs = @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass',
    '-File', $openScript,
    '-PartNumber', $pn,
    '-Open', 'FabricationZip'
  )
  if ($ForceZipOverwrite) { $openArgs += '-ForceZipOverwrite' }
  Invoke-NestedLauncherPowerShell -ArgumentList $openArgs -StepLabel 'Check and Build Fabrication Zip' | Out-Null
}

if ($UploadToArena) {
  Write-Host ''
  Write-Host '=== Upload Zip File to Arena (Gerbers) ==='
  $syncArgs = @('-File', $syncScript, '-PartNumber', $pn)
  if ($CreateZip -or $UseExistingFile) {
    $syncArgs += '-SkipZipBuildIfMissing'
  }
  if ($ForceUpload) {
    $syncArgs += '-ForceUpload'
  }
  $syncInvokeArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass') + $syncArgs
  Invoke-NestedLauncherPowerShell -ArgumentList $syncInvokeArgs -StepLabel 'Upload fabrication zip to Arena' | Out-Null
}

Write-Host ''
Write-Host 'Done.' -ForegroundColor DarkGreen
exit 0
