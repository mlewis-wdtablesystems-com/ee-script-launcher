<#
.SYNOPSIS
  Copy the latest WDTS PCB design checklist from the Altium Library folder into the part's main folder, then open it in Excel.

.DESCRIPTION
  1. Resolves the part's main design folder (same `partDir` as Open-AltiumPartFolder.ps1 -Open ResolveJson).
  2. Finds the newest template under the library folder whose file name (without extension) matches
     `WDTS - PCB Design Checklist vX.X` (version is dotted digits, e.g. v1.2 or v2.10.3).
     When multiple files match, the highest semantic version wins; ties break on last write time.
  3. Copies that file into the main folder as
     `{template base name} {PCBA} Rev {revision}{ext}`
     where PCBA is the canonical part number and revision comes from the production
     folder / files (ResolveJson revLetter). If revision cannot be determined, uses Rev A.
     Overwrites if that destination name already exists.
  4. Opens the copied file with the default application (Excel for .xlsx/.xlsm).

.PARAMETER PartNumber
  PPP-XX-XXX with optional folder suffix after a space (same rules as Open-AltiumPartFolder.ps1).

.PARAMETER LibraryFolder
  Folder containing the checklist templates. Default is the shared Altium Library path.
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [Parameter(Mandatory = $false)]
  [string]$LibraryFolder = 'C:\Users\mlewis\Walker Digital Table Systems\WTDS-RDAustralia - Documents\23. Altium\Altium Library'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$openScript = Join-Path $scriptsDir 'Open-AltiumPartFolder.ps1'
if (-not (Test-Path -LiteralPath $openScript)) { throw "Required script missing: $openScript" }

if (-not (Test-Path -LiteralPath $LibraryFolder -PathType Container)) {
  throw "Altium Library folder not found: $LibraryFolder"
}

Write-Host "=== Resolving main folder for $PartNumber ==="
$resolveOut = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File $openScript -PartNumber $PartNumber.Trim() -Open ResolveJson 2>&1
if ($LASTEXITCODE -ne 0) {
  throw "Open-AltiumPartFolder.ps1 -Open ResolveJson failed (exit $LASTEXITCODE):`n$($resolveOut -join [Environment]::NewLine)"
}
$joined = ($resolveOut | ForEach-Object { "$_" }) -join "`n"
$rxJson = [regex]'---BEGIN-RESOLVE-JSON---\s*(?<j>\{.*\})\s*---END-RESOLVE-JSON---'
$jm = $rxJson.Match($joined)
if (-not $jm.Success) {
  throw "Could not find resolve JSON markers. Output:`n$joined"
}
$resolved = $jm.Groups['j'].Value | ConvertFrom-Json
$partDir = $resolved.partDir
if ([string]::IsNullOrWhiteSpace($partDir)) {
  throw 'ResolveJson did not return partDir.'
}
if (-not (Test-Path -LiteralPath $partDir -PathType Container)) {
  throw "Resolved part folder does not exist: $partDir"
}
Write-Host "  Main folder: $partDir"

$pcba = ("$($resolved.canonical)").Trim()
if ([string]::IsNullOrWhiteSpace($pcba)) {
  # Fallback: strip optional folder suffix after a space (e.g. "272-10-001 BJ 5P").
  $pcba = ($PartNumber.Trim() -split '\s+', 2)[0].Trim()
}
$revRaw = ("$($resolved.revLetter)").Trim()
if ([string]::IsNullOrWhiteSpace($revRaw)) {
  $revLabel = 'A'
  Write-Host '  Revision: not detected from production folder/files; using Rev A'
} else {
  $revLabel = $revRaw.ToUpperInvariant()
  Write-Host ("  Revision: Rev {0} (from production folder/files)" -f $revLabel)
}
Write-Host ("  PCBA: {0}" -f $pcba)

Write-Host ''
Write-Host "=== Finding newest checklist under library ==="
$rxBase = [regex]'^(?i)WDTS - PCB Design Checklist v(?<ver>[\d.]+)$'
$candidates = New-Object System.Collections.Generic.List[object]
foreach ($f in @(Get-ChildItem -LiteralPath $LibraryFolder -File -ErrorAction SilentlyContinue)) {
  $m = $rxBase.Match($f.BaseName)
  if (-not $m.Success) { continue }
  $verStr = $m.Groups['ver'].Value
  $verObj = $null
  try { $verObj = [version]$verStr } catch { }
  $null = $candidates.Add([pscustomobject]@{
      File   = $f
      VerStr = $verStr
      VerObj = $verObj
    })
}
if ($candidates.Count -eq 0) {
  throw (
    "No file matching name pattern 'WDTS - PCB Design Checklist vX.X' was found under:`n  $LibraryFolder`n" +
    "Expected base name like 'WDTS - PCB Design Checklist v1.0' (any common extension)."
  )
}
$chosen = $candidates |
  Sort-Object @{ Expression = { if ($_.VerObj) { $_.VerObj } else { [version]'0.0' } }; Descending = $true },
  @{ Expression = { $_.File.LastWriteTimeUtc }; Descending = $true } |
  Select-Object -First 1
$src = $chosen.File
Write-Host ("  Using: {0} (version label v{1}, last write {2})" -f $src.Name, $chosen.VerStr, $src.LastWriteTime)

$destName = '{0} {1} Rev {2}{3}' -f $src.BaseName, $pcba, $revLabel, $src.Extension
$destPath = Join-Path $partDir $destName
Write-Host ''
Write-Host "=== Copying into part folder ==="
Write-Host "  Destination: $destPath"
Copy-Item -LiteralPath $src.FullName -Destination $destPath -Force

Write-Host ''
Write-Host '=== Opening in Excel ==='
Start-Process -FilePath $destPath
Write-Host 'Done.'
exit 0
