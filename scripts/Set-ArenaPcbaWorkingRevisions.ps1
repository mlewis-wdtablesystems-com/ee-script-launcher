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

.EXAMPLE
  .\Set-ArenaPcbaWorkingRevisions.ps1 -PartNumber 239-10-001 -Revision A1
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [Parameter(Mandatory = $true)] [string]$Revision
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

$revParsed = Parse-AltiumRevisionToken -RevisionText $Revision
if (-not $revParsed.Number) {
  throw "Revision must include a number after the letter (e.g. A1). Got: '$Revision'."
}

$parts = ConvertTo-PcbaAndBareBoardPartNumbers -PartNumber $PartNumber
$pcbaPn = $parts.Pcba
$barePn = $parts.Bare
$pcbaRev = $revParsed.Display
$bareRev = $revParsed.Letter

Write-Host '=== Set PCBA and PCB working revisions ===' -ForegroundColor Cyan
Write-Host ("  PCBA part:              {0}" -f $pcbaPn)
Write-Host ("  Bare board (Arena):     {0}" -f $barePn)
Write-Host ("  Revision (PCBA):        {0}" -f $pcbaRev)
Write-Host ("  Revision (bare letter): {0}" -f $bareRev)
Write-Host ''

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

Write-Host '=== Done ===' -ForegroundColor Green
exit 0
