<#
.SYNOPSIS
  Select and zoom to untented vias on the active PCB.

.DESCRIPTION
  Uses the WDTS Altium MCP bridge. Untented = not tented on Top or Bottom solder mask
  (IsTenting_Top / IsTenting_Bottom). Deselects all other objects, selects matching
  vias, zooms to the selection. If none are found, Altium shows "No untented vias".

  Requires Altium_API.PrjScr compiled once; active PCB document open.

.EXAMPLE
  .\Show-UntentedVias.ps1 -PartNumber 213-25-007
#>
param(
  [string]$PartNumber,
  [double]$TimeoutSec = 120
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$repoRoot = Join-Path $PSScriptRoot '..\wdts-altium-mcp'
$python = Join-Path $repoRoot '.venv\Scripts\python.exe'
$script = Join-Path $repoRoot 'scripts\show_untented_vias.py'

if (-not (Test-Path -LiteralPath $python)) {
  throw "WDTS MCP venv not found at $python. Run: cd wdts-altium-mcp; python -m venv .venv; pip install -e ."
}

$pyArgs = @('--timeout', $TimeoutSec)
if ($PartNumber) { $pyArgs += @('--part', $PartNumber) }
& $python $script @pyArgs
exit $LASTEXITCODE
