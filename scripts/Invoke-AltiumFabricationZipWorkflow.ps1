<#
.SYNOPSIS
  Optional fabrication zip build and/or Arena Gerber zip upload (launcher combined card).

.PARAMETER PartNumber
  Altium part lookup (PPP-XX-XXX).

.PARAMETER CreateZip
  Run Check and Build Fabrication Zip (Open-AltiumPartFolder -Open FabricationZip).

.PARAMETER UploadToArena
  Upload or replace "Altium Gerber Files" on the bare-board Arena working revision
  using the normal zip file name (FINAL).

.PARAMETER UploadAsDraft
  Same Arena upload as -UploadToArena, after renaming the zip so the file name
  ends in _DRAFT before the extension. Mutually exclusive with -UploadToArena.

  Leave -CreateZip off to upload a zip that is already in the production folder.
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [switch]$CreateZip,
  [switch]$UploadToArena,
  [switch]$UploadAsDraft,
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

if ($UploadToArena -and $UploadAsDraft) {
  throw 'Select only one Arena upload option: Upload Zip File to Arena FINAL or Upload to Arena as Draft.'
}
if (-not $CreateZip -and -not $UploadToArena -and -not $UploadAsDraft) {
  throw 'Enable at least one option: Create Zip, Upload Zip File to Arena FINAL, and/or Upload to Arena as Draft.'
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

if ($UploadToArena -or $UploadAsDraft) {
  Write-Host ''
  if ($UploadAsDraft) {
    Write-Host '=== Upload to Arena as Draft (Gerbers) ==='
  } else {
    Write-Host '=== Upload Zip File to Arena FINAL (Gerbers) ==='
  }
  # Create Zip already built the file. With Create Zip off, use the zip already
  # in the production folder and do not build one here.
  $syncArgs = @('-File', $syncScript, '-PartNumber', $pn, '-SkipZipBuildIfMissing')
  if ($UploadAsDraft) { $syncArgs += '-AsDraft' }
  if ($ForceUpload) { $syncArgs += '-ForceUpload' }
  $syncInvokeArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass') + $syncArgs
  $stepLabel = if ($UploadAsDraft) {
    'Upload fabrication zip to Arena as draft'
  } else {
    'Upload fabrication zip to Arena'
  }
  Invoke-NestedLauncherPowerShell -ArgumentList $syncInvokeArgs -StepLabel $stepLabel | Out-Null
}

Write-Host ''
Write-Host 'Done.' -ForegroundColor DarkGreen
exit 0
