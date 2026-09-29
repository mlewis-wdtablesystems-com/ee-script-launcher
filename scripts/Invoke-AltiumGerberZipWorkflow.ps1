<#
.SYNOPSIS
  Optional Gerber zip build (launcher combined card).

.PARAMETER PartNumber
  Altium part lookup (PPP-XX-XXX).

.PARAMETER CreateZip
  Run Check and Build Gerber Zip (Open-AltiumPartFolder -Open GerberZip).
  Zips Gerber and NC Drill only (no PFS) as "Gerbers PPP-YY-XXX Rev <letter>.zip".
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [switch]$CreateZip,
  [switch]$ForceZipOverwrite
)

$ErrorActionPreference = 'Stop'
$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $scriptsDir 'LauncherNestedScriptHelpers.ps1')
$openScript = Join-Path $scriptsDir 'Open-AltiumPartFolder.ps1'
if (-not (Test-Path -LiteralPath $openScript)) { throw "Required script missing: $openScript" }

if (-not $CreateZip) {
  throw 'Enable Create Zip to run this workflow.'
}

$pn = $PartNumber.Trim()

Write-Host '=== Create Zip (Check and Build Gerber Zip) ==='
$openArgs = @(
  '-NoProfile', '-ExecutionPolicy', 'Bypass',
  '-File', $openScript,
  '-PartNumber', $pn,
  '-Open', 'GerberZip'
)
if ($ForceZipOverwrite) { $openArgs += '-ForceZipOverwrite' }
Invoke-NestedLauncherPowerShell -ArgumentList $openArgs -StepLabel 'Check and Build Gerber Zip' | Out-Null

Write-Host ''
Write-Host 'Done.' -ForegroundColor DarkGreen
exit 0
