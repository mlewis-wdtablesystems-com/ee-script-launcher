<#
.SYNOPSIS
  Build fabrication zip if missing, then upload or replace "Altium Gerber Files" on the bare-board Arena item.

.DESCRIPTION
  End-to-end workflow:
    1. Resolve the Altium production folder via Open-AltiumPartFolder.ps1 -Open ResolveJson
       (same rules as Check and Build Fabrication Zip: newest rev with production outputs).
    2. Expected zip path:
         "<prodDoc>\Fabrication <PPP-20-YYY-or-25-YYY> Rev <letter>.zip"
       matching Check and Build Fabrication Zip naming (rev letter only, no trailing digit).
    3. If that zip does not exist, run Open-AltiumPartFolder.ps1 -Open FabricationZip
       (same checks and zip layout as the launcher card).
    4. Map the entered part number to the bare-board number (PPP-10→PPP-20, PPP-15→PPP-25;
       other middle segments pass through — use pcbPartNumber from resolve JSON).
    5. On Arena working revision for that bare-board item:
         - If no file in category "Altium Gerber Files": upload the zip as edition 1.
         - If one or more exist: add the next edition (current + 1) on the best-matching file
           (prefer title match, then highest edition). Unlocked rows are preferred when present;
           if all are locked, the same logic still runs so Replace-ArenaFile can post the new edition.

  Credentials: ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID.

.PARAMETER PartNumber
  Altium part lookup token (PPP-XX-XXX, optional descriptive tail). May be PCBA or bare board;
  the script uses ResolveJson's pcbPartNumber for Arena and zip naming.

.PARAMETER ForceUpload
  Skip the duplicate-filename confirmation (used when the launcher user chose Yes).

.PARAMETER AsDraft
  Rename the zip to add _DRAFT before the extension, then upload that file.
  If the normal zip is missing and the _DRAFT zip is already present, that file is uploaded.

.EXAMPLE
  .\Sync-ArenaFabricationGerbers.ps1 -PartNumber 294-10-022
  .\Sync-ArenaFabricationGerbers.ps1 -PartNumber 294-20-022
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [switch]$ForceUpload,
  [switch]$SkipZipBuildIfMissing,
  [switch]$AsDraft
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$openScript   = Join-Path $scriptsDir 'Open-AltiumPartFolder.ps1'
$uploadScript = Join-Path $scriptsDir 'Upload-ArenaFile.ps1'
$replaceScript = Join-Path $scriptsDir 'Replace-ArenaFile.ps1'
$getFilesScript = Join-Path $scriptsDir 'Get-ArenaItemFiles.ps1'
foreach ($p in @($openScript, $uploadScript, $replaceScript, $getFilesScript)) {
  if (-not (Test-Path -LiteralPath $p)) { throw "Required script missing: $p" }
}

$password = $env:ARENA_PASSWORD
if ([string]::IsNullOrWhiteSpace($env:ARENA_EMAIL)) { throw 'Set ARENA_EMAIL.' }
if ([string]::IsNullOrWhiteSpace($password))      { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_WORKSPACE_ID)) {
  throw 'Set ARENA_WORKSPACE_ID (launcher workspace dropdown sets this per run).'
}

function Get-ResolveJsonFromOpenScript {
  param([string]$Pn)
  $resolveOut = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
    -File $openScript -PartNumber $Pn -Open ResolveJson 2>&1
  if ($LASTEXITCODE -ne 0) {
    throw "Open-AltiumPartFolder.ps1 -Open ResolveJson failed (exit $LASTEXITCODE):`n$($resolveOut -join [Environment]::NewLine)"
  }
  $joined = ($resolveOut | ForEach-Object { "$_" }) -join "`n"
  $rxJson = [regex]'---BEGIN-RESOLVE-JSON---\s*(?<j>\{.*\})\s*---END-RESOLVE-JSON---'
  $jm = $rxJson.Match($joined)
  if (-not $jm.Success) {
    throw "Could not find resolve JSON markers. Output:`n$joined"
  }
  return ($jm.Groups['j'].Value | ConvertFrom-Json)
}

function Test-IsAltiumGerberCategory([object]$FileRow) {
  if (-not $FileRow) { return $false }
  $n = $FileRow.category_name
  if ($n -ne $null -and $n.ToString().Trim().ToLowerInvariant() -eq 'altium gerber files') {
    return $true
  }
  $path = $FileRow.category_path
  if ($path -ne $null -and $path.ToString().ToLowerInvariant().EndsWith('\altium gerber files')) {
    return $true
  }
  return $false
}

function Test-ArenaFileRowLocked([object]$row) {
  if (-not $row) { return $false }
  $v = $row.locked
  if ($null -eq $v) { return $false }
  if ($v -is [bool]) { return $v }
  $s = ("$v").Trim().ToLowerInvariant()
  return $s -in @('true', '1', 'yes')
}

function Test-ArenaDuplicateFileTitleOnRev {
  param(
    [object[]]$Files,
    [string[]]$TitlesToMatch,
    [string]$ExcludeFileGuid
  )
  $want = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach ($raw in $TitlesToMatch) {
    $t = ("$raw").Trim()
    if ($t.Length -gt 0) { [void]$want.Add($t) }
  }
  if ($want.Count -eq 0) { return $false }
  $ex = ("$ExcludeFileGuid").Trim()
  foreach ($f in $Files) {
    if (-not $f) { continue }
    $gid = ("$($f.file_guid)").Trim()
    if ($ex -and $gid -eq $ex) { continue }
    $tit = ("$($f.title)").Trim()
    if ($tit.Length -eq 0) { continue }
    if ($want.Contains($tit)) { return $true }
  }
  return $false
}

function Rename-ProductionZipAsDraft {
  param([Parameter(Mandatory = $true)][string]$ZipPath)
  $leaf = Split-Path -Leaf $ZipPath
  if ($leaf -match '(?i)_DRAFT\.zip$') { return $ZipPath }
  if ($leaf -notmatch '(?i)\.zip$') {
    throw "Expected a .zip file to rename as draft: $ZipPath"
  }
  $draftLeaf = $leaf -replace '(?i)\.zip$', '_DRAFT.zip'
  $dest = Join-Path (Split-Path -Parent $ZipPath) $draftLeaf
  if (Test-Path -LiteralPath $dest) {
    Remove-Item -LiteralPath $dest -Force
  }
  Rename-Item -LiteralPath $ZipPath -NewName $draftLeaf
  Write-Host "  Renamed zip for draft upload: $dest"
  return $dest
}

# ---------- 1) Altium resolve ----------
Write-Host "=== Altium resolve ==="
$resolved = Get-ResolveJsonFromOpenScript -Pn $PartNumber.Trim()
$prodDocDir = $resolved.prodDocDir
if ([string]::IsNullOrWhiteSpace($prodDocDir)) {
  throw 'ResolveJson did not return prodDocDir (no production folder under any revision?).'
}
$barePn = $resolved.pcbPartNumber
if ([string]::IsNullOrWhiteSpace($barePn)) {
  throw 'ResolveJson did not return pcbPartNumber.'
}

Write-Host "  Production document folder: $prodDocDir"
Write-Host "  Bare-board part (Arena):    $barePn"

$revLetter = $resolved.revLetter
if ([string]::IsNullOrWhiteSpace($revLetter)) {
  $revLetter = 'A'
  Write-Host '  Revision (from Altium):     (not detected; assuming Rev A)' -ForegroundColor DarkYellow
}
else {
  Write-Host "  Revision (from Altium):     $revLetter"
}

$revLetterOnly = ($revLetter -replace '\d+$', '')
if ([string]::IsNullOrWhiteSpace($revLetterOnly)) { $revLetterOnly = $revLetter }
$zipBase = "Fabrication $barePn Rev $revLetterOnly"
$zipPath = Join-Path $prodDocDir ($zipBase + '.zip')
$draftPath = Join-Path $prodDocDir ($zipBase + '_DRAFT.zip')

Write-Host "  Expected zip:               $zipPath"
if ($AsDraft) {
  Write-Host "  Draft zip name:             $draftPath"
}

# ---------- 2) Build zip if missing ----------
if (-not (Test-Path -LiteralPath $zipPath)) {
  if ($AsDraft -and (Test-Path -LiteralPath $draftPath)) {
    Write-Host "  Using existing draft zip:   $draftPath"
    $zipPath = $draftPath
  }
  elseif ($SkipZipBuildIfMissing) {
    throw "Fabrication zip not found (Create Zip was not selected): $zipPath"
  }
  else {
    Write-Host ''
    Write-Host "=== Fabrication zip not found; running Check and Build Fabrication Zip ==="
    $fabOut = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
      -File $openScript -PartNumber $PartNumber.Trim() -Open FabricationZip 2>&1
    $fabText = ($fabOut | ForEach-Object { "$_" }) -join "`n"
    Write-Host $fabText
    if ($LASTEXITCODE -ne 0) {
      throw "Check and Build Fabrication Zip failed (exit $LASTEXITCODE)."
    }
  }
}

if (-not (Test-Path -LiteralPath $zipPath)) {
  throw "Fabrication zip still missing after build: $zipPath"
}

if ($AsDraft) {
  $zipPath = Rename-ProductionZipAsDraft -ZipPath $zipPath
}

$zipItem = Get-Item -LiteralPath $zipPath
Write-Host ''
Write-Host ("=== Using zip ({0} bytes) ===" -f $zipItem.Length)

# ---------- 3) Arena — list working-rev files ----------
Write-Host ''
Write-Host "=== Arena: files on working revision ($barePn) ==="
$filesRaw = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File $getFilesScript -PartNumber $barePn -Revision working 2>&1
if ($LASTEXITCODE -ne 0) {
  throw "Get-ArenaItemFiles.ps1 failed (exit $LASTEXITCODE):`n$($filesRaw -join [Environment]::NewLine)"
}
$filesPayload = ($filesRaw | Out-String).Trim() | ConvertFrom-Json
$allFiles = @($filesPayload.files)
$gerbers = @($allFiles | Where-Object { Test-IsAltiumGerberCategory $_ })

$revInfo = $filesPayload.rev
if ($revInfo) {
  Write-Host ("  Arena working revision: number='{0}'  guid={1}" -f $revInfo.number, $revInfo.guid)
}

Write-Host ("  Total files on working rev: {0}" -f $allFiles.Count)
Write-Host ("  'Altium Gerber Files' files: {0}" -f $gerbers.Count)

$titleForUpload = [System.IO.Path]::GetFileNameWithoutExtension($zipItem.Name)
$titlesForDup = @($zipItem.Name, $titleForUpload)

$dupExcludeGuid = $null
$best = $null
if ($gerbers.Count -gt 0) {
  $unlocked = @($gerbers | Where-Object { -not (Test-ArenaFileRowLocked $_) })
  $pool = $unlocked
  if ($pool.Count -eq 0) {
    Write-Host (
      "Arena file list marked all $($gerbers.Count) 'Altium Gerber Files' row(s) as locked; " +
      "using best title/edition match for the next edition anyway."
    ) -ForegroundColor DarkGray
    $pool = $gerbers
  }
  $pref = @($pool | Where-Object {
    $t = ("$($_.title)").Trim()
    $t -eq $titleForUpload -or $t -eq $zipItem.Name
  })
  $usePool = if ($pref.Count -gt 0) { $pref } else { $pool }
  $best = @(
    $usePool | Sort-Object @{
      Expression = {
        $ed = 0
        [void][int]::TryParse(("$($_.edition)").Trim(), [ref]$ed)
        $ed
      }; Descending = $true
    }
  ) | Select-Object -First 1
  if (-not $best) { $best = $pool[0] }
  $dupExcludeGuid = ("$($best.file_guid)").Trim()
}

if (Test-ArenaDuplicateFileTitleOnRev -Files $allFiles -TitlesToMatch $titlesForDup -ExcludeFileGuid $dupExcludeGuid) {
  if (-not $ForceUpload) {
    Write-Host '---DUPLICATE-FILENAME-PROMPT---'
    [pscustomobject]@{
      duplicate = $true
      message   = 'File with this File Name Already Exists, Upload Anyway?'
    } | ConvertTo-Json -Compress
    exit 10
  }
}

if ($gerbers.Count -eq 0) {
  Write-Host ''
  Write-Host "=== Upload new Altium Gerber Files (edition 1) ==="
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $uploadScript `
    -PartNumber $barePn `
    -FilePath $zipPath `
    -Title $titleForUpload `
    -Edition '1' `
    -CategoryName 'Altium Gerber Files' `
    -Revision working
  if ($LASTEXITCODE -ne 0) { throw "Upload-ArenaFile.ps1 failed (exit $LASTEXITCODE)." }
  Write-Host "Done. Uploaded new Gerber file for $barePn."
  exit 0
}

Write-Host ''
Write-Host ("=== Replace existing Gerber file {0} (edition {1}) ===" -f $best.number, $best.edition)
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $replaceScript `
  -FileGuid $best.file_guid `
  -FilePath $zipPath `
  -CurrentEdition ("$($best.edition)".Trim())
if ($LASTEXITCODE -ne 0) { throw "Replace-ArenaFile.ps1 failed (exit $LASTEXITCODE)." }

Write-Host "Done. New edition uploaded for $($best.number) on $barePn."
exit 0
