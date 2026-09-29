<#
.SYNOPSIS
  Upload working-revision files to Arena for a PCBA (not including BOM).

.DESCRIPTION
  For a PCBA part number (PPP-10-XXX / PPP-15-XXX) and revision (letter + number, e.g. A1):

    1. Fabrication zip — upload existing zip to bare-board Arena item (no zip build).
    2. Pick & Place zip (optional) — build zip, upload, use existing if present; revision must match Altium letter.
    3. Sync Arena Product Image on the PCBA item (skipped if an image at that revision already exists).
    4. Set PCBA working revision number to the entered revision (e.g. A1).
    5. Set bare-board (PPP-20/25-XXX) working revision to the letter only (e.g. A).
    6. Run Check PCBA and PCB Files.

  Credentials: ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID.

.PARAMETER PartNumber
  PCBA part number for Altium lookup and Arena (e.g. 239-10-001).

.PARAMETER Revision
  Two-character style revision: one letter + one number (e.g. A1 or A 1).

.PARAMETER IncludePickPlace
  When set, run Pick & Place zip (build, upload, use existing). Omit to skip that step.

.EXAMPLE
  .\Upload-AllArenaWorkingRevisionFiles.ps1 -PartNumber 239-10-001 -Revision A1
.EXAMPLE
  .\Upload-AllArenaWorkingRevisionFiles.ps1 -PartNumber 239-10-001 -Revision A1 -IncludePickPlace
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [Parameter(Mandatory = $true)] [string]$Revision,
  [switch]$IncludePickPlace
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $scriptsDir 'AltiumRevisionHelpers.ps1')

$fabWorkflow   = Join-Path $scriptsDir 'Invoke-AltiumFabricationZipWorkflow.ps1'
$ppWorkflow    = Join-Path $scriptsDir 'Invoke-AltiumPickPlaceZipWorkflow.ps1'
$imageSync     = Join-Path $scriptsDir 'Sync-ArenaProductImage.ps1'
$setRevScript  = Join-Path $scriptsDir 'Set-ArenaItemWorkingRevisionNumber.ps1'
$checkScript   = Join-Path $scriptsDir 'Check-ArenaPcbaFiles.ps1'
$requiredScripts = @($fabWorkflow, $imageSync, $setRevScript, $checkScript)
if ($IncludePickPlace) { $requiredScripts += $ppWorkflow }
foreach ($p in $requiredScripts) {
  if (-not (Test-Path -LiteralPath $p)) { throw "Required script missing: $p" }
}

if ([string]::IsNullOrWhiteSpace($env:ARENA_EMAIL)) { throw 'Set ARENA_EMAIL.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_PASSWORD)) { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_WORKSPACE_ID)) {
  throw 'Set ARENA_WORKSPACE_ID (launcher workspace dropdown sets this per run).'
}

$revParsed = Parse-AltiumRevisionToken -RevisionText $Revision
if (-not $revParsed.Number) {
  throw "Revision must include a number after the letter (e.g. A1). Got: '$Revision'."
}

$parts = ConvertTo-PcbaAndBareBoardPartNumbers -PartNumber $PartNumber
$pcbaPn = $parts.Pcba
$barePn = $parts.Bare
$pcbaRev = $revParsed.Display
$bareRev = $revParsed.Letter

Write-Host '=== Upload Fab, P&P, Image Files ===' -ForegroundColor Cyan
Write-Host ("  PCBA part:              {0}" -f $pcbaPn)
Write-Host ("  Bare board (Arena):     {0}" -f $barePn)
Write-Host ("  Revision (PCBA):        {0}" -f $pcbaRev)
Write-Host ("  Revision (bare letter): {0}" -f $bareRev)
Write-Host ("  Include Pick & Place:   {0}" -f $(if ($IncludePickPlace) { 'yes' } else { 'no' }))
Write-Host ''

function Invoke-Step {
  param(
    [Parameter(Mandatory = $true)] [string]$Title,
    [Parameter(Mandatory = $true)] [scriptblock]$Action
  )
  Write-Host ''
  Write-Host "=== $Title ===" -ForegroundColor Cyan
  & $Action
  if ($LASTEXITCODE -ne 0) {
    throw "Step failed (exit $LASTEXITCODE): $Title"
  }
}

Invoke-Step -Title 'Fabrication zip (upload existing file only)' -Action {
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $fabWorkflow `
    -PartNumber $pcbaPn -UploadToArena -UseExistingFile
}

if ($IncludePickPlace) {
  Invoke-Step -Title 'Pick & Place zip (build, upload, use existing)' -Action {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $ppWorkflow `
      -PartNumber $pcbaPn `
      -RevisionOverride $pcbaRev `
      -CreateZip -UploadToArena -UseExistingFile
  }
} else {
  Write-Host ''
  Write-Host '=== Pick & Place zip ===' -ForegroundColor Cyan
  Write-Host '  Skipped (Include Pick & Place not selected).' -ForegroundColor DarkGray
}

Invoke-Step -Title 'Sync Arena Product Image' -Action {
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $imageSync `
    -PartNumber $pcbaPn
}

Invoke-Step -Title "Set PCBA working revision ($pcbaPn -> $pcbaRev)" -Action {
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $setRevScript `
    -PartNumber $pcbaPn -RevisionNumber $pcbaRev
}

Invoke-Step -Title "Set bare-board working revision ($barePn -> $bareRev)" -Action {
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $setRevScript `
    -PartNumber $barePn -RevisionNumber $bareRev
}

Invoke-Step -Title 'Check PCBA and PCB Files' -Action {
  $checkArgs = @(
    '-PartNumber', $pcbaPn,
    '-ExpectedRevision', $pcbaRev
  )
  if (-not $IncludePickPlace) { $checkArgs += '-SkipPickPlace' }
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $checkScript @checkArgs
}

Write-Host ''
Write-Host '=== All steps completed ===' -ForegroundColor DarkGreen
exit 0
