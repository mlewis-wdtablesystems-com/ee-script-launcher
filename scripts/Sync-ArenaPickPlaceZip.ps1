<#
.SYNOPSIS
  Build Pick & Place zip if missing, then upload or replace "Pick and Place" on the PCBA Arena item.

.DESCRIPTION
  End-to-end workflow:
    1. Resolve Altium production folder via Open-AltiumPartFolder.ps1 -Open ResolveJson.
    2. Expected zip path (same as Build Pick & Place zip):
         "<prodDoc>\Pick Place <canonical> Rev <fullRev>.zip"
       where canonical is PPP-XX-XXX from the resolve JSON and fullRev includes any trailing digit (e.g. B1).
    3. If that zip does not exist, run Open-AltiumPartFolder.ps1 -Open PickPlaceZip.
    4. Derive the PCBA / assembly part number for Arena:
         PPP-20-YYY -> PPP-10-YYY, PPP-25-YYY -> PPP-15-YYY; PPP-10/15 unchanged.
    5. On Arena working revision for that assembly item (Get-ArenaItemFiles -Revision working):
         - If no "Pick and Place" category file: upload the zip as edition 1.
         - Otherwise replace an existing file: prefer rows whose title matches this zip, then
           highest edition. If every row is locked, the same chain is still used so the next
           edition (current + 1) can be posted; see Replace-ArenaFile.ps1 if Arena rejects.

  Credentials: ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID.

.PARAMETER PartNumber
  Altium lookup (PPP-XX-XXX, optional tail). May be PCBA or bare board; zip name follows
  ResolveJson canonical; Arena always uses the mapped PPP-10 / PPP-15 part.

.PARAMETER ForceUpload
  Skip the duplicate-filename confirmation (used when the launcher user chose Yes).

.EXAMPLE
  .\Sync-ArenaPickPlaceZip.ps1 -PartNumber 294-10-022
  .\Sync-ArenaPickPlaceZip.ps1 -PartNumber 294-20-022
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [switch]$ForceUpload,
  [switch]$SkipZipBuildIfMissing
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$openScript    = Join-Path $scriptsDir 'Open-AltiumPartFolder.ps1'
$uploadScript  = Join-Path $scriptsDir 'Upload-ArenaFile.ps1'
$replaceScript = Join-Path $scriptsDir 'Replace-ArenaFile.ps1'
$getFilesScript = Join-Path $scriptsDir 'Get-ArenaItemFiles.ps1'
foreach ($p in @($openScript, $uploadScript, $replaceScript, $getFilesScript)) {
  if (-not (Test-Path -LiteralPath $p)) { throw "Required script missing: $p" }
}

if ([string]::IsNullOrWhiteSpace($env:ARENA_EMAIL)) { throw 'Set ARENA_EMAIL.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_PASSWORD)) { throw 'Set ARENA_PASSWORD.' }
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

function Resolve-AssemblyPartNumber([string]$canonical) {
  $segs = $canonical.Trim().Split('-')
  if ($segs.Count -ne 3) { return $canonical.Trim() }
  $mid = $segs[1]
  $asm = switch ($mid) {
    '20' { '10' }
    '25' { '15' }
    default { $mid }
  }
  return ('{0}-{1}-{2}' -f $segs[0], $asm, $segs[2])
}

function Test-IsPickPlaceCategory([object]$FileRow) {
  if (-not $FileRow) { return $false }
  $n = $FileRow.category_name
  if ($n -ne $null -and $n.ToString().Trim().ToLowerInvariant() -eq 'pick and place') {
    return $true
  }
  $path = $FileRow.category_path
  if ($path -ne $null -and $path.ToString().ToLowerInvariant().EndsWith('\pick and place')) {
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

# ---------- 1) Altium resolve ----------
Write-Host "=== Altium resolve ==="
$resolved = Get-ResolveJsonFromOpenScript -Pn $PartNumber.Trim()
$prodDocDir = $resolved.prodDocDir
if ([string]::IsNullOrWhiteSpace($prodDocDir)) {
  throw 'ResolveJson did not return prodDocDir (no production folder under any revision?).'
}
$canonical = $resolved.canonical
if ([string]::IsNullOrWhiteSpace($canonical)) {
  throw 'ResolveJson did not return canonical.'
}
$revLetter = $resolved.revLetter
if ([string]::IsNullOrWhiteSpace($revLetter)) {
  throw 'ResolveJson did not return revLetter; cannot name Pick & Place zip.'
}

$assemblyPn = Resolve-AssemblyPartNumber $canonical
$asmSegs = $assemblyPn.Split('-')
if ($asmSegs.Count -eq 3) {
  $m = $asmSegs[1]
  if ($m -ne '10' -and $m -ne '15') {
    Write-Warning "Mapped Arena part number '$assemblyPn' is not PPP-10-YYY or PPP-15-YYY (middle segment '$m'). Upload may target the wrong item."
  }
}

Write-Host "  Production document folder: $prodDocDir"
Write-Host "  Canonical (zip name):       $canonical"
Write-Host "  PCBA part (Arena):          $assemblyPn"
Write-Host "  Revision (from Altium):     $revLetter"

$zipBase = "Pick Place $canonical Rev $revLetter"
$zipPath = Join-Path $prodDocDir ($zipBase + '.zip')
Write-Host "  Expected zip:               $zipPath"

# ---------- 2) Build zip if missing ----------
if (-not (Test-Path -LiteralPath $zipPath)) {
  if ($SkipZipBuildIfMissing) {
    throw "Pick & Place zip not found (upload-only / use existing file): $zipPath"
  }
  Write-Host ''
  Write-Host "=== Pick & Place zip not found; running Build Pick & Place zip ==="
  $ppOut = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
    -File $openScript -PartNumber $PartNumber.Trim() -Open PickPlaceZip 2>&1
  Write-Host (($ppOut | ForEach-Object { "$_" }) -join "`n")
  if ($LASTEXITCODE -ne 0) {
    throw "Build Pick & Place zip failed (exit $LASTEXITCODE)."
  }
}

if (-not (Test-Path -LiteralPath $zipPath)) {
  throw "Pick & Place zip still missing after build: $zipPath"
}

$zipItem = Get-Item -LiteralPath $zipPath
Write-Host ''
Write-Host ("=== Using zip ({0} bytes) ===" -f $zipItem.Length)

# ---------- 3) Arena ----------
Write-Host ''
Write-Host "=== Arena: files on working revision ($assemblyPn) ==="
$filesRaw = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File $getFilesScript -PartNumber $assemblyPn -Revision working 2>&1
if ($LASTEXITCODE -ne 0) {
  throw "Get-ArenaItemFiles.ps1 failed (exit $LASTEXITCODE):`n$($filesRaw -join [Environment]::NewLine)"
}
$filesPayload = ($filesRaw | Out-String).Trim() | ConvertFrom-Json
$allFiles = @($filesPayload.files)
$pnp = @($allFiles | Where-Object { Test-IsPickPlaceCategory $_ })

$revInfo = $filesPayload.rev
if ($revInfo) {
  Write-Host ("  Arena working revision: number='{0}'  guid={1}" -f $revInfo.number, $revInfo.guid)
}

Write-Host ("  Total files on working rev: {0}" -f $allFiles.Count)
Write-Host ("  'Pick and Place' files:     {0}" -f $pnp.Count)

$titleForUpload = [System.IO.Path]::GetFileNameWithoutExtension($zipItem.Name)
$titlesForDup = @($zipItem.Name, $titleForUpload)

$dupExcludeGuid = $null
$best = $null
if ($pnp.Count -gt 0) {
  $pnpUnlocked = @($pnp | Where-Object { -not (Test-ArenaFileRowLocked $_) })
  $pool = $pnpUnlocked
  if ($pool.Count -eq 0) {
    Write-Warning (
      "All $($pnp.Count) 'Pick and Place' file(s) on the working revision report locked/checked out. " +
      "Selecting the best-matching file by title and edition to post the next edition (current + 1) anyway."
    )
    $pool = $pnp
  }

  # Prefer rows whose title matches this zip (avoid picking a stale / other Pick&Place chain).
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

if ($pnp.Count -eq 0) {
  Write-Host ''
  Write-Host "=== Upload new Pick and Place (edition 1) ==="
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $uploadScript `
    -PartNumber $assemblyPn `
    -FilePath $zipPath `
    -Title $titleForUpload `
    -Edition '1' `
    -CategoryName 'Pick and Place' `
    -Revision working
  if ($LASTEXITCODE -ne 0) { throw "Upload-ArenaFile.ps1 failed (exit $LASTEXITCODE)." }
  Write-Host "Done. Uploaded new Pick and Place file for $assemblyPn."
  exit 0
}

Write-Host ''
Write-Host ("=== Replace existing Pick and Place {0} (edition {1}) ===" -f $best.number, $best.edition)
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $replaceScript `
  -FileGuid $best.file_guid `
  -FilePath $zipPath `
  -CurrentEdition ("$($best.edition)".Trim())
if ($LASTEXITCODE -ne 0) { throw "Replace-ArenaFile.ps1 failed (exit $LASTEXITCODE)." }

Write-Host "Done. New edition uploaded for $($best.number) on $assemblyPn."
exit 0
