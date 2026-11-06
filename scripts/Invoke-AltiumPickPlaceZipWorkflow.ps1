<#
.SYNOPSIS
  Optional Pick & Place zip build and/or Arena upload (launcher combined card).

.PARAMETER PartNumber
  Altium part lookup (PPP-XX-XXX).

.PARAMETER CreateZip
  Run Build Pick & Place zip (Open-AltiumPartFolder -Open PickPlaceZip).

.PARAMETER UploadToArena
  Upload or replace "Pick and Place" on the PCBA Arena working revision
  using the normal zip file name (FINAL).

.PARAMETER UploadAsDraft
  Same Arena upload as -UploadToArena, after renaming the zip so the file name
  ends in _DRAFT before the extension. Mutually exclusive with -UploadToArena.

.PARAMETER RevisionOverride
  Optional revision for the zip name (letter + number, e.g. D1). The letter must match the
  revision letter from Altium outputs; the number you specify is used in the zip file name.

  Leave -CreateZip off to upload a zip that is already in the production folder.
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [string]$RevisionOverride = '',
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
$syncScript = Join-Path $scriptsDir 'Sync-ArenaPickPlaceZip.ps1'
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

if ($UploadToArena -or $UploadAsDraft) {
  Write-Host ''
  if ($UploadAsDraft) {
    Write-Host '=== Upload to Arena as Draft (Pick & Place) ==='
  } else {
    Write-Host '=== Upload Zip File to Arena FINAL (Pick & Place) ==='
  }
  $syncArgs = @('-File', $syncScript, '-PartNumber', $pn, '-SkipZipBuildIfMissing')
  if (-not [string]::IsNullOrWhiteSpace($RevisionOverride)) {
    $syncArgs += '-RevisionOverride', $RevisionOverride.Trim()
  }
  if ($UploadAsDraft) { $syncArgs += '-AsDraft' }
  if ($ForceUpload) { $syncArgs += '-ForceUpload' }
  $syncInvokeArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass') + $syncArgs
  $stepLabel = if ($UploadAsDraft) {
    'Upload Pick & Place zip to Arena as draft'
  } else {
    'Upload Pick & Place zip to Arena'
  }
  Invoke-NestedLauncherPowerShell -ArgumentList $syncInvokeArgs -StepLabel $stepLabel | Out-Null
}

Write-Host ''
Write-Host 'Done.' -ForegroundColor DarkGreen
exit 0
