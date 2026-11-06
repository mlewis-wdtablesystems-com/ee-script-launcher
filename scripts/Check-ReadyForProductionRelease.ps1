<#
.SYNOPSIS
  Run pre-production release checks for a PCBA on working revisions.

.DESCRIPTION
  For a PCBA part number, runs (in order) and prints the full output of each.
  PCB checks use the standard PPP-20/25 mapping unless that part is missing
  from the PCBA BOM and another PCB is present; then that BOM PCB is used.

    1. Remove Illustrative Image from PCB bare board (working revision), if present
    2. Arena checked-in revision letter vs Altium production folder
    3. Check Pick and Place File with Arena BOM (working revision)
    4. Check PCBA and PCB Files (working revision). Fails if any file attached
       to the PCBA or PCB has _DRAFT in the file name.
    5. Compare Working and Checked In Revision

  All steps run even when an earlier step fails. Exits non-zero if any step failed.

.PARAMETER PartNumber
  PCBA part number (e.g. 251-10-001 or 251-15-001).

.EXAMPLE
  .\Check-ReadyForProductionRelease.ps1 -PartNumber 251-10-001
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $scriptsDir 'LauncherNestedScriptHelpers.ps1')

$removePcbImageScript = Join-Path $scriptsDir 'Remove-ArenaPcbIllustrativeImageAssociations.ps1'
$revMatchScript  = Join-Path $scriptsDir 'Check-ArenaCheckedInRevMatchesAltiumProduction.ps1'
$pickPlaceScript = Join-Path $scriptsDir 'Check-PickPlaceFileWithArenaBom.ps1'
$pcbaFilesScript = Join-Path $scriptsDir 'Check-ArenaPcbaFiles.ps1'
$compareScript   = Join-Path $scriptsDir 'Compare-WorkingAndCheckedInRevision.ps1'
foreach ($p in @($removePcbImageScript, $revMatchScript, $pickPlaceScript, $pcbaFilesScript, $compareScript)) {
  if (-not (Test-Path -LiteralPath $p)) { throw "Required script missing: $p" }
}

if ([string]::IsNullOrWhiteSpace($env:ARENA_EMAIL)) { throw 'Set ARENA_EMAIL.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_PASSWORD)) { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_WORKSPACE_ID)) {
  throw 'Set ARENA_WORKSPACE_ID (launcher workspace dropdown sets this per run).'
}

$pn = $PartNumber.Trim()

Write-Host '=== Check Ready for Production Release ===' -ForegroundColor Cyan
Write-Host ("  Part number: {0}" -f $pn)
Write-Host ''

function Invoke-ReleaseCheckStep {
  param(
    [Parameter(Mandatory = $true)] [string]$Title,
    [Parameter(Mandatory = $true)] [string[]]$ArgumentList
  )

  Write-Host ''
  Write-Host "=== $Title ===" -ForegroundColor Cyan

  $prev = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'
  try {
    $out = & powershell.exe @ArgumentList 2>&1
    $code = $LASTEXITCODE
  }
  finally {
    $ErrorActionPreference = $prev
  }

  $text = Format-NestedPowerShellOutput $out
  if ($text.Trim()) {
    Write-ScriptResultLines $text
  }

  if ($code -ne 0) {
    $summary = Get-NestedScriptFailureSummary -OutputText $text -StepLabel $Title -ExitCode $code
    Write-Host ''
    Write-ScriptResultLines "FAILED: $summary"
  } else {
    Write-ScriptResultLines ("  OK: {0}" -f $Title)
  }

  return [pscustomobject]@{
    Title    = $Title
    ExitCode = $code
  }
}

$psBase = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File')

$results = New-Object System.Collections.Generic.List[object]
foreach ($step in @(
    @{ Title = 'Remove Illustrative Image from PCB (bare board)'; Args = @($removePcbImageScript, '-PartNumber', $pn) }
    @{ Title = 'Arena checked-in revision letter vs Altium production folder'; Args = @($revMatchScript, '-PartNumber', $pn) }
    @{ Title = 'Check Pick and Place File with Arena BOM (working revision)'; Args = @($pickPlaceScript, '-PartNumber', $pn, '-RevisionMode', 'working') }
    @{ Title = 'Check PCBA and PCB Files (working revision)'; Args = @($pcbaFilesScript, '-PartNumber', $pn, '-RevisionMode', 'working') }
    @{ Title = 'Compare Working and Checked In Revision'; Args = @($compareScript, '-PartNumber', $pn) }
  )) {
  $one = Invoke-ReleaseCheckStep -Title $step.Title -ArgumentList ($psBase + $step.Args)
  [void]$results.Add($one)
}

$failed = @($results | Where-Object { $null -ne $_.ExitCode -and $_.ExitCode -ne 0 })
Write-Host ''
if ($failed.Count -gt 0) {
  Write-ScriptResultLines ("=== Summary: {0} of {1} checks failed ===" -f $failed.Count, $results.Count)
  foreach ($step in $failed) {
    Write-ScriptResultLines ("  - {0} (exit {1})" -f $step.Title, $step.ExitCode)
  }
  exit 1
}

Write-ScriptResultLines '=== All checks passed ==='
exit 0
