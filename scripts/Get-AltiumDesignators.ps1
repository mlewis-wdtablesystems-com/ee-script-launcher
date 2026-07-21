<#
.SYNOPSIS
  List PCB designators for a WDTS part via the Altium MCP bridge.

.DESCRIPTION
  Resolves the part to .PcbDoc / .PrjPcb, runs get_all_component_data on the live
  bridge, and prints sorted designators. Falls back to the newest Pick & Place CSV
  under the part folder if the bridge times out.

.EXAMPLE
  .\Get-AltiumDesignators.ps1 -PartNumber 213-15-007

.EXAMPLE
  .\Get-AltiumDesignators.ps1 -PartNumber 213-25-007
#>
param(
  [Parameter(Mandatory = $true)]
  [string]$PartNumber,
  [double]$TimeoutSec = 300
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$repoRoot = Join-Path $PSScriptRoot '..\wdts-altium-mcp'
$python = Join-Path $repoRoot '.venv\Scripts\python.exe'
$script = Join-Path $repoRoot 'scripts\get_designators.py'

if (-not (Test-Path -LiteralPath $python)) {
  throw "WDTS MCP venv not found at $python. Run: cd wdts-altium-mcp; python -m venv .venv; pip install -e ."
}

& $python $script $PartNumber --timeout $TimeoutSec
exit $LASTEXITCODE
