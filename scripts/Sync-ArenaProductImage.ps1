<#
.SYNOPSIS
  Upload or replace an Illustrative Image on the PCBA Arena item from Altium "PCB 3D Print" outputs.

.DESCRIPTION
  End-to-end workflow:
    1. Resolve Altium production folder via Open-AltiumPartFolder.ps1 -Open ResolveJson.
    2. Look for .png files in "<prodDocDir>\PCB 3D Print" (non-recursive). If several exist,
       the newest by LastWriteTimeUtc is used (with a warning). If the PNG file name contains
       a bare-board part number (PPP-20-NNN or PPP-25-NNN), rename it on disk to the PCBA
       equivalent (PPP-10-NNN or PPP-15-NNN) before upload.
    3. Derive the PCBA part number for Arena (same mapping as Pick & Place):
         PPP-20-YYY -> PPP-10-YYY, PPP-25-YYY -> PPP-15-YYY; PPP-10/15 unchanged.
    4. Arena file title: "Product Image <assemblyPn> Rev <L>" where <L> is letters-only from
       ResolveJson revLetter (e.g. B1 -> B, A -> A).
    5. File category: "Illustrative Image".
    6. On the working revision for that assembly item:
         - If no Illustrative Image: upload as edition 1.
         - Otherwise: new edition on the existing file with the highest numeric edition
           (includes locked rows if none are unlocked; see Replace-ArenaFile.ps1 if Arena rejects).

  Credentials: ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID.

.PARAMETER PartNumber
  Altium lookup (PPP-XX-XXX, optional tail). May be PCBA or bare board; image folder follows
  ResolveJson prodDocDir; Arena always uses the mapped PPP-10 / PPP-15 part.

.PARAMETER ForceUpload
  Skip the duplicate-filename confirmation (used when the launcher user chose Yes).

.EXAMPLE
  .\Sync-ArenaProductImage.ps1 -PartNumber 294-10-022
  .\Sync-ArenaProductImage.ps1 -PartNumber 294-20-022
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [switch]$ForceUpload
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$openScript     = Join-Path $scriptsDir 'Open-AltiumPartFolder.ps1'
$uploadScript   = Join-Path $scriptsDir 'Upload-ArenaFile.ps1'
$replaceScript  = Join-Path $scriptsDir 'Replace-ArenaFile.ps1'
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

function Convert-BareBoardPartNumbersInFileName([string]$FileName) {
  # PPP-20-NNN -> PPP-10-NNN, PPP-25-NNN -> PPP-15-NNN (case-insensitive; leave other text intact).
  if ([string]::IsNullOrWhiteSpace($FileName)) { return $FileName }
  $out = [regex]::Replace($FileName, '(?i)\b(\d{3})-20-(\d{3})\b', '$1-10-$2')
  return [regex]::Replace($out, '(?i)\b(\d{3})-25-(\d{3})\b', '$1-15-$2')
}

function Rename-ProductImagePngIfBareBoardNamed {
  param([Parameter(Mandatory = $true)] [System.IO.FileInfo]$PngItem)
  $oldName = $PngItem.Name
  $newName = Convert-BareBoardPartNumbersInFileName $oldName
  if ($newName -ceq $oldName) { return $PngItem }

  $dest = Join-Path $PngItem.DirectoryName $newName
  if (
    (Test-Path -LiteralPath $dest) -and
    -not [string]::Equals($PngItem.FullName, $dest, [StringComparison]::OrdinalIgnoreCase)
  ) {
    throw (
      "Cannot rename product image PNG from '$oldName' to '$newName': " +
      "destination already exists at `"$dest`"."
    )
  }

  Write-Host ("  Renaming PNG (bare board PN -> PCBA): '{0}' -> '{1}'" -f $oldName, $newName)
  Rename-Item -LiteralPath $PngItem.FullName -NewName $newName -Force
  return (Get-Item -LiteralPath $dest)
}

function Get-RevLetterOnly([string]$revLetter) {
  if ([string]::IsNullOrWhiteSpace($revLetter)) { return '' }
  $m = [regex]::Match($revLetter.Trim(), '^[A-Za-z]+')
  if ($m.Success) { return $m.Value.ToUpperInvariant() }
  throw "Could not derive a letter-only revision from '$revLetter'."
}

function Get-RevLetterFromTitle([string]$Title) {
  if ([string]::IsNullOrWhiteSpace($Title)) { return $null }
  $m = [regex]::Match($Title.Trim(), '(?i)\bRev\s+([A-Za-z]+)')
  if ($m.Success) { return $m.Groups[1].Value.ToUpperInvariant() }
  return $null
}

function Test-IllustrativeImageAlreadyAtRevision {
  param(
    [object[]]$IllustrativeRows,
    [string]$ExpectedTitle,
    [string]$RevLetterOnly
  )
  foreach ($row in $IllustrativeRows) {
    if (-not $row) { continue }
    $t = ("$($row.title)").Trim()
    if ($t.Length -eq 0) { continue }
    if ($t -ceq $ExpectedTitle) { return $true }
    $fileRev = Get-RevLetterFromTitle -Title $t
    if ($fileRev -and $fileRev -ceq $RevLetterOnly) { return $true }
  }
  return $false
}

function Test-IsIllustrativeImageCategory([object]$FileRow) {
  if (-not $FileRow) { return $false }
  $n = $FileRow.category_name
  if ($null -ne $n -and $n.ToString().Trim().ToLowerInvariant() -eq 'illustrative image') {
    return $true
  }
  $path = $FileRow.category_path
  if ($null -ne $path -and $path.ToString().ToLowerInvariant().EndsWith('\illustrative image')) {
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
$revLetterRaw = $resolved.revLetter
if ([string]::IsNullOrWhiteSpace($revLetterRaw)) {
  throw 'ResolveJson did not return revLetter; cannot name the Arena image title.'
}

$assemblyPn = Resolve-AssemblyPartNumber $canonical
$asmSegs = $assemblyPn.Split('-')
if ($asmSegs.Count -eq 3) {
  $m = $asmSegs[1]
  if ($m -ne '10' -and $m -ne '15') {
    Write-Warning "Mapped Arena part number '$assemblyPn' is not PPP-10-YYY or PPP-15-YYY (middle segment '$m'). Upload may target the wrong item."
  }
}

$revLetterOnly = Get-RevLetterOnly $revLetterRaw
$arenaTitle = "Product Image $assemblyPn Rev $revLetterOnly"

$pcb3dDir = Join-Path $prodDocDir 'PCB 3D Print'
Write-Host "  Production document folder: $prodDocDir"
Write-Host "  PCB 3D Print folder:        $pcb3dDir"
Write-Host "  Canonical:                  $canonical"
Write-Host "  PCBA part (Arena):          $assemblyPn"
Write-Host "  Revision (letter only):     $revLetterOnly"
Write-Host "  Arena file title:           $arenaTitle"

if (-not (Test-Path -LiteralPath $pcb3dDir -PathType Container)) {
  throw "PCB 3D Print folder not found: $pcb3dDir"
}

$pngFiles = @(Get-ChildItem -LiteralPath $pcb3dDir -File -Filter '*.png' -ErrorAction SilentlyContinue)
if ($pngFiles.Count -eq 0) {
  throw "No .png file found in `"$pcb3dDir`"."
}
if ($pngFiles.Count -gt 1) {
  Write-Warning "Multiple .png files in PCB 3D Print; using the newest by last write time."
}
$pngItem = $pngFiles | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
$pngItem = Rename-ProductImagePngIfBareBoardNamed -PngItem $pngItem
$pngPath = $pngItem.FullName

Write-Host ''
Write-Host ("=== Using PNG ({0} bytes): {1} ===" -f $pngItem.Length, $pngItem.Name)

# ---------- 2) Arena ----------
Write-Host ''
Write-Host "=== Arena: files on working revision ($assemblyPn) ==="
$filesRaw = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File $getFilesScript -PartNumber $assemblyPn -Revision working 2>&1
if ($LASTEXITCODE -ne 0) {
  throw "Get-ArenaItemFiles.ps1 failed (exit $LASTEXITCODE):`n$($filesRaw -join [Environment]::NewLine)"
}
$filesPayload = ($filesRaw | Out-String).Trim() | ConvertFrom-Json
$allFiles = @($filesPayload.files)
$illus = @($allFiles | Where-Object { Test-IsIllustrativeImageCategory $_ })

$revInfo = $filesPayload.rev
if ($revInfo) {
  Write-Host ("  Arena working revision: number='{0}'  guid={1}" -f $revInfo.number, $revInfo.guid)
}

Write-Host ("  Total files on working rev: {0}" -f $allFiles.Count)
Write-Host ("  'Illustrative Image' files: {0}" -f $illus.Count)

if (Test-IllustrativeImageAlreadyAtRevision -IllustrativeRows $illus -ExpectedTitle $arenaTitle -RevLetterOnly $revLetterOnly) {
  Write-Host ''
  Write-Host '=== Skipping product image upload ===' -ForegroundColor DarkGray
  Write-Host (
    "  Illustrative Image already on working revision for Rev $revLetterOnly " +
    "(title matches '$arenaTitle' or same revision letter)."
  )
  exit 0
}

$titlesForDup = @(
  $pngItem.Name
  [System.IO.Path]::GetFileNameWithoutExtension($pngItem.Name)
  $arenaTitle
)

$dupExcludeGuid = $null
$best = $null
if ($illus.Count -gt 0) {
  $unlocked = @($illus | Where-Object { -not (Test-ArenaFileRowLocked $_) })
  $pool = $unlocked
  if ($pool.Count -eq 0) {
    Write-Host (
      "Arena file list marked all $($illus.Count) 'Illustrative Image' row(s) as locked; " +
      "using best title/edition match for the next edition anyway."
    ) -ForegroundColor DarkGray
    $pool = $illus
  }
  $pref = @($pool | Where-Object {
    $t = ("$($_.title)").Trim()
    $t -eq $arenaTitle
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

if ($illus.Count -eq 0) {
  Write-Host ''
  Write-Host "=== Upload new Illustrative Image (edition 1) ==="
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $uploadScript `
    -PartNumber $assemblyPn `
    -FilePath $pngPath `
    -Title $arenaTitle `
    -Edition '1' `
    -CategoryName 'Illustrative Image' `
    -Revision working
  if ($LASTEXITCODE -ne 0) { throw "Upload-ArenaFile.ps1 failed (exit $LASTEXITCODE)." }
  Write-Host "Done. Uploaded new Illustrative Image for $assemblyPn."
  exit 0
}

Write-Host ''
Write-Host ("=== Replace existing Illustrative Image {0} (edition {1}) ===" -f $best.number, $best.edition)
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $replaceScript `
  -FileGuid $best.file_guid `
  -FilePath $pngPath `
  -CurrentEdition ("$($best.edition)".Trim()) `
  -Title $arenaTitle
if ($LASTEXITCODE -ne 0) { throw "Replace-ArenaFile.ps1 failed (exit $LASTEXITCODE)." }

Write-Host "Done. New edition uploaded for $($best.number) on $assemblyPn."
exit 0
