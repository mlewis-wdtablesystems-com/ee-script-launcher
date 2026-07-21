<#
.SYNOPSIS
  Check Top Overlay silkscreen for bare PCB and assembly part numbers + revision.

.DESCRIPTION
  Uses the WDTS Altium MCP bridge on the active PCB. Assembly part number (PPP-XX-XXX)
  is taken from the focused .PrjPcb file name. Bare PCB part number follows Script
  Launcher rules (PPP-10-XXX -> PPP-20-XXX, PPP-15-XXX -> PPP-25-XXX). Revision is
  parsed from Rev/Ver in the project or PCB document name.

  Requires Altium_API.PrjScr compiled once and the active PCB + project open.

.EXAMPLE
  .\Check-PartNumberSilkscreen.ps1 -Part 213-15-007

.EXAMPLE
  .\Check-PartNumberSilkscreen.ps1 -Part 213-25-007
#>
param(
  [string]$Part,
  [double]$TimeoutSec = 180
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$repoRoot = Join-Path $PSScriptRoot '..\wdts-altium-mcp'
$python = Join-Path $repoRoot '.venv\Scripts\python.exe'
$script = Join-Path $repoRoot 'scripts\check_part_number_silkscreen.py'

if (-not (Test-Path -LiteralPath $python)) {
  throw "WDTS MCP venv not found at $python. Run: cd wdts-altium-mcp; python -m venv .venv; pip install -e ."
}

$pyArgs = @('--timeout', $TimeoutSec)
if ($Part) { $pyArgs += @('--part', $Part) }
& $python $script @pyArgs
exit $LASTEXITCODE
