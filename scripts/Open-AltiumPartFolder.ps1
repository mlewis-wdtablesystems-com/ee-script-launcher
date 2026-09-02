<#
.SYNOPSIS
  Resolves the Altium project folder for a top-level part number (PPP-XX-XXX), optionally
  followed by descriptive text matching the filesystem folder name, picks the latest
  revision subfolder when present (Rev A, Rev B, Ver …), then opens Explorer or the newest *.PrjPcb in Altium.

.DESCRIPTION
  Searches under the Walker Digital Table Systems "23. Altium" root for directories whose
  name equals PPP-XX-XXX or whose name begins with PPP-XX-XXX followed by a space,
  underscore, or hyphen (e.g. "272-10-001 BJ 5P ...", "251-15-001_Ver A 20260716",
  "258-10-001-ASY 1D Cage ANT"). A trailing date after Rev/Ver is allowed. You can pass:
    • PPP-XX-XXX only — when multiple matching folders exist, pick the one with the newest
      *.PcbDoc inside (excluding Archive); if none have a PCB, fall back to folder mtime
    • PPP-XX-XXX plus a suffix after a space — if that exact folder name exists under a parent, it wins

  With -Open Explorer (default), opens that folder (or the latest Rev / Ver subfolder) in Explorer.
  With -Open ProductionExplorer, picks the newest revision that has production outputs, then opens the
  production document folder ("Production Document*", "Project Outputs for *", or a direct child with
  Gerber/BOM) in Explorer — same resolution as fabrication/pick-place zip flows.
  With -Open PrjPcb, opens the most recently modified *.PrjPcb under that same resolved folder
  (recursive, excluding paths under a folder named "Archive").
  With -Open PcbDoc, opens the most recently modified *.PcbDoc under that same resolved folder
  (same search rules as PrjPcb).

  Locations tried:
    • {Root}\* matching part folder name rule
    • {Root}\PPP\* matching rule (PPP = first hyphen segment)

  Scan: every other top-level directory for a subdirectory matching the rule.

  If the resolved part folder contains revision subfolders (Rev …, Ver …), the script picks
  the latest by suffix ordering (numbers numerically max; letters ordinal desc with mtime tie-break)
  before opening Explorer or searching for .PrjPcb files. A folder named like
  "Project Outputs for … Cage Ver A1" is **not** treated as a Rev/Ver folder (it is the output bundle);
  Gerber/BOM under it resolve as the production folder even when there is no extra nesting.

.PARAMETER PartNumber
  PPP-XX-XXX, optionally followed by space and additional text mirroring the folder name (e.g. "272-10-001 BJ 5P").

.PARAMETER Open
  Explorer — open the resolved design folder in Windows Explorer (default).
  ProductionExplorer — open the production document folder (same discovery as FabricationZip /
  PickPlaceZip: newest rev with outputs, then "Production Document*", "Project Outputs for *", or
  child with Gerber/BOM).
  PrjPcb — after resolving the same folder (including latest Rev/Ver), search recursively for
  *.PrjPcb files (excluding any path under a folder named "Archive") and open the most recently
  modified file with the default application (Altium).
  PcbDoc — same folder resolution and Archive exclusion, but opens the newest *.PcbDoc instead.
  FabricationZip — go one level under the resolved Rev/Ver folder, locate the production document
  folder ("Production Document", "Project Outputs for *", or any direct child containing a "Gerber"
  subfolder), inspect the Gerber folder for files sharing an extension, for Mechanical Layer
  plots other than Layer 1 (*.GM2+; orange warnings if any), and for missing required layer
  extensions (GM, GM1, GTP, GTL, GTS; orange warnings if any), then zip "Gerber", "NC Drill", and
  "PFS" into "Fabrication PPP-YY-XXX Rev ?.zip" inside that production document folder.
  GerberZip — same production-document resolution and folder checks as FabricationZip (Gerber and
  NC Drill only; PFS is not required or included), then zip "Gerber" and "NC Drill" into
  "Gerbers PPP-YY-XXX Rev ?.zip" inside that production document folder.
  PickPlaceZip — same production-document resolution, then zip the "Pick Place" folder into
  "Pick Place PPP-TT-XXX Rev ?.zip" inside the production document folder. Errors if the folder
  contains more than one *.csv or more than one *.txt file (top-level). When -RevisionOverride
  is set (e.g. D1 from the launcher), that value names the zip; Altium folder/PrjPcb revision is
  used only for an optional letter mismatch check when detectable.

  PPP-YY-XXX maps TT→YY for the assembly→PCB part number scheme (TT=10→YY=20, TT=15→YY=25; other
  values pass through unchanged). The revision letter is read from the resolved folder name (Rev/Ver
  suffix) or, as a fallback, from any *.PrjPcb file name in the resolved folder.

.PARAMETER Root
  Root directory containing Altium project folders. Default:
  C:\Users\mlewis\Walker Digital Table Systems\WTDS-RDAustralia - Documents\23. Altium
#>
param(
  [Parameter(Mandatory = $true)]
  [string]$PartNumber,
  [Parameter(Mandatory = $false)]
  [ValidateSet('Explorer', 'ProductionExplorer', 'PrjPcb', 'PcbDoc', 'FabricationZip', 'GerberZip', 'PickPlaceZip', 'CheckFabrication', 'ResolveJson')]
  [string]$Open = 'Explorer',
  [Parameter(Mandatory = $false)]
  [int]$TimestampSpreadMinutes = 2,
  [Parameter(Mandatory = $false)]
  [string]$Root = 'C:\Users\mlewis\Walker Digital Table Systems\WTDS-RDAustralia - Documents\23. Altium',
  [Parameter(Mandatory = $false)]
  [switch]$ForceZipOverwrite,
  [Parameter(Mandatory = $false)]
  [string]$RevisionOverride = ''
)

$ErrorActionPreference = 'Stop'
$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $scriptsDir 'AltiumRevisionHelpers.ps1')
. (Join-Path $scriptsDir 'AltiumZipHelpers.ps1')
. (Join-Path $scriptsDir 'ExplorerWindowHelpers.ps1')

function Get-AltiumSearchRoots([string]$RootParam) {
  $r = if ([string]::IsNullOrWhiteSpace($RootParam)) {
    'C:\Users\mlewis\Walker Digital Table Systems\WTDS-RDAustralia - Documents\23. Altium'
  } else {
    $RootParam.Trim()
  }
  if (-not (Test-Path -LiteralPath $r)) {
    throw "Altium root folder not found: $r"
  }
  return @((Resolve-Path -LiteralPath $r).Path)
}

function Invoke-ZipOverwritePromptIfNeeded {
  param(
    [Parameter(Mandatory = $true)] [string]$ZipPath,
    [switch]$ForceZipOverwrite
  )
  if ($ForceZipOverwrite) { return }
  if (-not (Test-Path -LiteralPath $ZipPath)) { return }
  $item = Get-Item -LiteralPath $ZipPath
  Write-Host '---ZIP-OVERWRITE-PROMPT---'
  [pscustomobject]@{
    zipOverwrite  = $true
    message     = "Existing zip will be overwritten: $($item.Name)"
    zipPath     = $item.FullName
    zipName     = $item.Name
    lastWrite   = $item.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
    lengthBytes = $item.Length
  } | ConvertTo-Json -Compress
  exit 11
}

function Split-AltiumPartArg([string]$raw) {
  $t = $raw.Trim()
  if (-not $t) { throw 'Part number is empty.' }
  # Core token PPP-XX-XXX; optional descriptive suffix starting after whitespace
  $m = [regex]::Match($t, '^(?<core>\d{3}-\d{2}-\d{3})(?:\s+(?<tail>.+))?$')
  if (-not $m.Success) {
    throw "Expected PPP-XX-XXX with optional descriptive suffix after a space (e.g. `"272-10-001`" or `"272-10-001 BJ 5P`")."
  }
  $core = $m.Groups['core'].Value
  $tail = if ($m.Groups['tail'].Success) { $m.Groups['tail'].Value.Trim() } else { '' }
  return @{ Canonical = $core; NameTail = $tail }
}

function Test-PartFolderTitleMatch([string]$dirName, [string]$canonical) {
  # Exact name, or canonical then space/underscore/hyphen + rest
  # (e.g. "251-15-001 BJ 5P", "251-15-001_Ver A 20260716", "258-10-001-ASY 1D Cage ANT").
  ($dirName -eq $canonical) -or ($dirName -match ('^' + [regex]::Escape($canonical) + '([\s_\-]|$)'))
}

function Add-PathToList([System.Collections.Generic.List[string]]$list, [string]$candidatePath) {
  if (-not (Test-Path -LiteralPath $candidatePath)) { return }
  $full = (Resolve-Path -LiteralPath $candidatePath).Path
  if ($list -notcontains $full) { $null = $list.Add($full) }
}

function Find-PartFoldersUnderParent(
  [string]$ParentDir,
  [string]$Canonical,
  [string]$OptionalNameTail,
  [System.Collections.Generic.List[string]]$OutCandidates
) {
  if (-not (Test-Path -LiteralPath $ParentDir)) { return }

  $children = @(Get-ChildItem -LiteralPath $ParentDir -Directory -ErrorAction SilentlyContinue)

  if ($OptionalNameTail) {
    $exactName = "{0} {1}" -f $Canonical, $OptionalNameTail
    foreach ($ch in $children) {
      if ($ch.Name -eq $exactName) {
        Add-PathToList $OutCandidates $ch.FullName
        return
      }
    }
  }

  foreach ($ch in $children) {
    if (Test-PartFolderTitleMatch $ch.Name $Canonical) {
      Add-PathToList $OutCandidates $ch.FullName
    }
  }
}

function Parse-RevSuffix([string]$dirName) {
  # Altium "Project Outputs for …" / "Production Document…" folders often end
  # with text like "Cage Ver A1". That is not an assembly revision folder; it
  # is the production output bundle itself. Do not treat it as Rev/Ver.
  if ($dirName -imatch '^\s*Project Outputs for\s' -or $dirName -ilike 'Production Document*') {
    return $null
  }
  # Accept any folder name that contains a rev tag:
  #   "Ver A", "Rev C1", "Revision B", "VerB2",
  #   "Some Product Rev A", "Daughter Board Ver C1",
  #   "251-15-001_Ver A 20260716" (underscore before Ver; optional date/trailing text).
  # The rev keyword must be at the start OR preceded by whitespace/_/-,
  # and the alphanumeric suffix (letters then optional digits) may be followed
  # by more text (e.g. a date).
  if ($dirName -match '(?i)(?:^|[\s_\-])(?:Rev(?:ision)?|Ver\.?)\s*[:\s_\-]*(?<suffix>[A-Z]+\d*)(?:\b|_|$)') {
    return ($matches['suffix'].Trim().ToUpper())
  }
  return $null
}

function Get-RevLetterFromName([string]$name) {
  # Capture the rev "letter" as one or more letters optionally followed by
  # digits, e.g. "A", "B3", "AA", "C1". Folder naming can be inconsistent
  # (e.g. "Ver C1" vs "Ver C"), and we want both to resolve cleanly.
  # The trailing (?:\b|_|$) tolerates "Ver C1", "Ver C1_foo", "Ver C foo".
  if ($name -match '(?i)(?:^|[\s_\-])Rev(?:ision)?\s*[:\s_\-]*([A-Z]+\d*)(?:\b|_|$)') {
    return $matches[1].ToUpper()
  }
  if ($name -match '(?i)(?:^|[\s_\-])Ver\.?\s*[:\s_\-]*([A-Z]+\d*)(?:\b|_|$)') {
    return $matches[1].ToUpper()
  }
  return $null
}

function Write-OrangeWarning([string]$message) {
  $supportsVt = $false
  try { $supportsVt = [bool]$Host.UI.SupportsVirtualTerminal } catch { $supportsVt = $false }

  if ($supportsVt) {
    $esc = [char]27
    $open = "$esc[38;2;255;165;0m"
    $close = "$esc[0m"
    Write-Host ("{0}{1}{2}" -f $open, $message, $close)
  }
  else {
    Write-Host $message -ForegroundColor DarkYellow
  }
}

function Write-RedError([string]$message) {
  $supportsVt = $false
  try { $supportsVt = [bool]$Host.UI.SupportsVirtualTerminal } catch { $supportsVt = $false }

  if ($supportsVt) {
    $esc = [char]27
    $open = "$esc[38;2;255;0;0m"
    $close = "$esc[0m"
    Write-Host ("{0}{1}{2}" -f $open, $message, $close)
  }
  else {
    Write-Host $message -ForegroundColor Red
  }
}

function Test-GerberFolderForDuplicateExtensions([string]$gerberDir) {
  if (-not (Test-Path -LiteralPath $gerberDir -PathType Container)) { return @() }

  $files = @(Get-ChildItem -LiteralPath $gerberDir -File -ErrorAction SilentlyContinue)
  if ($files.Count -lt 2) { return @() }

  $warnings = New-Object System.Collections.Generic.List[string]
  $duplicateGroups = $files |
    Group-Object -Property { $_.Extension.ToLower() } |
    Where-Object { $_.Count -gt 1 }

  foreach ($grp in $duplicateGroups) {
    $extLabel = if ([string]::IsNullOrEmpty($grp.Name)) { '(no extension)' } else { $grp.Name }
    $names = ($grp.Group | Sort-Object Name | ForEach-Object { $_.Name }) -join ', '
    $null = $warnings.Add(('WARNING: Multiple {0} files in Gerber folder: {1}' -f $extLabel, $names))
  }
  return @($warnings)
}

function Test-GerberFolderForExtraMechanicalLayers([string]$gerberDir) {
  # Altium plots Mechanical Layer N as *.GMn (e.g. .GM1, .GM2, .GM16).
  # Only Mechanical Layer 1 is normally expected in fabrication Gerbers.
  if (-not (Test-Path -LiteralPath $gerberDir -PathType Container)) { return @() }

  $files = @(Get-ChildItem -LiteralPath $gerberDir -File -ErrorAction SilentlyContinue)
  $extra = @()
  foreach ($f in $files) {
    $m = [regex]::Match($f.Extension, '^\.GM(\d+)$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $m.Success) { continue }
    $layerNum = [int]$m.Groups[1].Value
    if ($layerNum -ne 1) {
      $extra += [pscustomobject]@{ Layer = $layerNum; Name = $f.Name }
    }
  }
  if ($extra.Count -eq 0) { return @() }

  $names = ($extra | Sort-Object Layer, Name | ForEach-Object { $_.Name }) -join ', '
  $layers = ($extra | Sort-Object Layer -Unique | ForEach-Object { $_.Layer }) -join ', '
  return @(
    ('WARNING: Gerber folder contains Mechanical Layer file(s) other than Layer 1 (layers {0}): {1}' -f $layers, $names)
  )
}

function Test-GerberFolderForRequiredExtensions([string]$gerberDir) {
  # Warn when expected copper/mask/paste/mechanical layer extensions are absent.
  $required = @('GM', 'GM1', 'GTP', 'GTL', 'GTS')
  if (-not (Test-Path -LiteralPath $gerberDir -PathType Container)) {
    return @(
      ('WARNING: Gerber folder missing required file extension(s): {0} (Gerber folder not found).' -f ($required -join ', '))
    )
  }

  $present = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
  foreach ($f in @(Get-ChildItem -LiteralPath $gerberDir -File -ErrorAction SilentlyContinue)) {
    $ext = $f.Extension
    if ([string]::IsNullOrEmpty($ext)) { continue }
    [void]$present.Add($ext.TrimStart('.'))
  }

  $missing = @($required | Where-Object { -not $present.Contains($_) })
  if ($missing.Count -eq 0) { return @() }
  return @(
    ('WARNING: Gerber folder missing required file extension(s): {0}.' -f ($missing -join ', '))
  )
}

function Test-FolderTimestampSpread {
  # All files inside $folder must have LastWriteTime within $maxSpreadMinutes
  # of each other. Returns a single error string if not, or $null if OK /
  # not applicable (folder missing or <2 files).
  param(
    [string]$Folder,
    [string]$Label,
    [int]$MaxSpreadMinutes
  )
  if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { return $null }
  $files = @(Get-ChildItem -LiteralPath $Folder -File -ErrorAction SilentlyContinue)
  if ($files.Count -lt 2) { return $null }
  $newest = ($files | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1)
  $stale = @($files | Where-Object {
      ($newest.LastWriteTimeUtc - $_.LastWriteTimeUtc).TotalMinutes -gt $MaxSpreadMinutes
    } | Sort-Object LastWriteTimeUtc)
  if ($stale.Count -eq 0) { return $null }
  $oldest = $stale[0]
  $spread = ($newest.LastWriteTimeUtc - $oldest.LastWriteTimeUtc)
  $spreadDisplay = if ($spread.TotalMinutes -lt 60) {
    "{0:N1} minutes" -f $spread.TotalMinutes
  } elseif ($spread.TotalHours -lt 48) {
    "{0:N1} hours" -f $spread.TotalHours
  } else {
    "{0:N1} days" -f $spread.TotalDays
  }
  $staleList = ($stale | ForEach-Object {
      ("''{0}'' @ {1:yyyy-MM-dd HH:mm}" -f $_.Name, $_.LastWriteTime)
    }) -join '; '
  return ("ERROR: $Label folder file timestamps span $spreadDisplay (limit $MaxSpreadMinutes min). " +
    ('Newest: ''{0}'' @ {1:yyyy-MM-dd HH:mm}. ' -f $newest.Name, $newest.LastWriteTime) +
    "Stale file(s) not from the same export (delete leftover orphans or re-run outputs): $staleList.")
}

function Invoke-FabricationFolderChecks {
  # Runs all the fabrication folder checks. Returns:
  #   @{ Errors = [string[]]; Warnings = [string[]] }
  param(
    [string]$ProdDocDir,
    [int]$TimestampSpreadMinutes = 2,
    [switch]$IncludePfsChecks,
    [switch]$AllowTwoNcDrillFiles
  )
  if (-not $PSBoundParameters.ContainsKey('IncludePfsChecks')) {
    $IncludePfsChecks = $true
  }
  $errors   = New-Object System.Collections.Generic.List[string]
  $warnings = New-Object System.Collections.Generic.List[string]

  $pfsDir = Join-Path $ProdDocDir 'PFS'

  # 1. PFS must contain exactly one .pdf (Fabrication zip only)
  if ($IncludePfsChecks) {
    $pfsPdfs = @()
    if (Test-Path -LiteralPath $pfsDir -PathType Container) {
      $pfsPdfs = @(Get-ChildItem -LiteralPath $pfsDir -File -Filter '*.pdf' -ErrorAction SilentlyContinue)
    }
    if ($pfsPdfs.Count -ne 1) {
      $listed = if ($pfsPdfs.Count -eq 0) { '(none)' } else { ($pfsPdfs | Sort-Object Name | ForEach-Object { $_.Name }) -join ', ' }
      $null = $errors.Add("ERROR: PFS folder must contain exactly one .pdf file (found $($pfsPdfs.Count): $listed).")
    }
  }

  # 2. Gerber must contain at least 8 files
  $gerberDir = Join-Path $ProdDocDir 'Gerber'
  $gerberFiles = @()
  if (Test-Path -LiteralPath $gerberDir -PathType Container) {
    $gerberFiles = @(Get-ChildItem -LiteralPath $gerberDir -File -ErrorAction SilentlyContinue)
  }
  if ($gerberFiles.Count -lt 8) {
    $null = $errors.Add("ERROR: Gerber folder must contain at least 8 files (found $($gerberFiles.Count)).")
  }

  # 3. NC Drill minimum depends on Gerber file count (simple boards: <=10 Gerber -> min 2)
  $ncDrillMin = if ($gerberFiles.Count -le 10) { 2 } else { 3 }
  $ncDrillDir = Join-Path $ProdDocDir 'NC Drill'
  $ncDrillFiles = @()
  if (Test-Path -LiteralPath $ncDrillDir -PathType Container) {
    $ncDrillFiles = @(Get-ChildItem -LiteralPath $ncDrillDir -File -ErrorAction SilentlyContinue)
  }
  if ($ncDrillFiles.Count -lt $ncDrillMin) {
    $listed = if ($ncDrillFiles.Count -eq 0) { '(none)' } else { ($ncDrillFiles | Sort-Object Name | ForEach-Object { $_.Name }) -join ', ' }
    $gerberNote = if ($gerberFiles.Count -le 10) {
      "Gerber has $($gerberFiles.Count) file(s) (<=10), minimum $ncDrillMin NC Drill file(s) required."
    } else {
      "Gerber has $($gerberFiles.Count) file(s) (>10), minimum $ncDrillMin NC Drill file(s) required."
    }
    $ncDrillMsg = "NC Drill folder must contain at least $ncDrillMin file(s) (found $($ncDrillFiles.Count): $listed). $gerberNote"
    if ($AllowTwoNcDrillFiles -and $ncDrillFiles.Count -eq 2) {
      $null = $warnings.Add("WARNING: Only 2 NC Drill file(s) ($listed). $gerberNote Acceptable for zip build.")
    }
    else {
      $null = $errors.Add("ERROR: $ncDrillMsg")
    }
  }

  # 4. Gerber duplicate-extension warnings
  foreach ($w in Test-GerberFolderForDuplicateExtensions $gerberDir) { $null = $warnings.Add($w) }

  # 4b. Extra mechanical layers (*.GM2, *.GM3, …) — warn; *.GM1 is OK
  foreach ($w in Test-GerberFolderForExtraMechanicalLayers $gerberDir) { $null = $warnings.Add($w) }

  # 4c. Required layer extensions (warn if missing)
  foreach ($w in Test-GerberFolderForRequiredExtensions $gerberDir) { $null = $warnings.Add($w) }

  # 5. Per-folder timestamp consistency (must be within TimestampSpreadMinutes)
  $timestampFolders = @(
    @{ Path = $gerberDir;  Label = 'Gerber'   }
    @{ Path = $ncDrillDir; Label = 'NC Drill' }
  )
  if ($IncludePfsChecks) {
    $timestampFolders += @{ Path = $pfsDir; Label = 'PFS' }
  }
  foreach ($pair in $timestampFolders) {
    $e = Test-FolderTimestampSpread -Folder $pair.Path -Label $pair.Label -MaxSpreadMinutes $TimestampSpreadMinutes
    if ($e) { $null = $errors.Add($e) }
  }

  return @{ Errors = @($errors); Warnings = @($warnings) }
}

function Show-FabricationFolderCheckReport {
  # Same printed layout as CheckFabrication — shared so FabricationZip shows it too.
  param(
    [string]$ProdDocDir,
    [string]$RevLetter,
    [int]$TimestampSpreadMinutes,
    [hashtable]$CheckResult,
    [string]$ReportTitle = 'Fabrication folder check'
  )
  $errs = @($CheckResult.Errors)
  $warns = @($CheckResult.Warnings)
  Write-Host ''
  Write-Host "--- $ReportTitle ---"
  Write-Host "Production document folder: $ProdDocDir"
  Write-Host "Revision: $RevLetter"
  Write-Host "Timestamp spread limit: $TimestampSpreadMinutes minute(s)"
  Write-Host ''
  if ($warns.Count -gt 0) {
    foreach ($w in $warns) { Write-OrangeWarning $w }
    Write-Host ''
  }
  if ($errs.Count -gt 0) {
    foreach ($e in $errs) { Write-RedError $e }
    Write-Host ''
  }
  if ($errs.Count -eq 0 -and $warns.Count -eq 0) {
    Write-Host 'All checks passed.'
  }
  else {
    Write-Host ("Summary: {0} error(s), {1} warning(s)." -f $errs.Count, $warns.Count)
  }
}

function Resolve-ProductionDocFolder([string]$targetDir) {
  if (-not (Test-Path -LiteralPath $targetDir -PathType Container)) {
    throw "Production document resolve: path not found: $targetDir"
  }

  $leaf = Split-Path -Leaf $targetDir
  # Some layouts place Gerber/BOM directly under "Project Outputs for …"
  # (no extra nesting). If we already sit in that folder, use it.
  if ($leaf -ilike 'Production Document*' -or $leaf -imatch '^\s*Project Outputs for\s') {
    foreach ($probe in @('Gerber', 'BOM', 'NC Drill', 'PFS', 'Pick Place')) {
      if (Test-Path -LiteralPath (Join-Path $targetDir $probe) -PathType Container) {
        return $targetDir
      }
    }
  }

  $children = @(Get-ChildItem -LiteralPath $targetDir -Directory -ErrorAction SilentlyContinue)

  # Accept any folder name starting with "Production Document" (covers
  # "Production Document" and "Production Documentation"), or starting with
  # "Project Outputs for ".
  foreach ($d in $children) {
    if ($d.Name -ilike 'Production Document*' -or $d.Name -ilike 'Project Outputs for *') {
      return $d.FullName
    }
  }

  foreach ($d in $children) {
    $gerberCandidate = Join-Path $d.FullName 'Gerber'
    if (Test-Path -LiteralPath $gerberCandidate -PathType Container) {
      return $d.FullName
    }
    $bomCandidate = Join-Path $d.FullName 'BOM'
    if (Test-Path -LiteralPath $bomCandidate -PathType Container) {
      return $d.FullName
    }
  }

  throw 'Production Document Folder Not Found'
}

function Test-AltiumPathContainsFolderSegment([string]$fullPath, [string]$folderSegment) {
  if ([string]::IsNullOrWhiteSpace($fullPath) -or [string]::IsNullOrWhiteSpace($folderSegment)) { return $false }
  foreach ($seg in ($fullPath -split '[\\/]')) {
    if ($seg -and ($seg -ieq $folderSegment)) { return $true }
  }
  return $false
}

function Get-AltiumPrjPcbFilesExcludingArchive([string]$rootDir) {
  $all = @(Get-ChildItem -LiteralPath $rootDir -Recurse -File -Filter '*.PrjPcb' -ErrorAction SilentlyContinue)
  return @($all | Where-Object { -not (Test-AltiumPathContainsFolderSegment $_.FullName 'Archive') })
}

function Get-AltiumPcbDocFilesExcludingArchive([string]$rootDir) {
  $all = @(Get-ChildItem -LiteralPath $rootDir -Recurse -File -Filter '*.PcbDoc' -ErrorAction SilentlyContinue)
  return @($all | Where-Object { -not (Test-AltiumPathContainsFolderSegment $_.FullName 'Archive') })
}

function Select-RevFolderWithProductionDoc {
  # Newest revision folder that actually resolves to a production-document area.
  # Matches ResolveJson behaviour: avoids picking an empty newer rev (e.g. only
  # Firmware) when Gerber/NC Drill/PFS still live under an older rev.
  param(
    [Parameter(Mandatory)][string]$PartFolder,
    [AllowEmptyCollection()][object[]]$RevRecords
  )
  $revs = @($RevRecords)
  if ($revs.Count -eq 0) {
    return $PartFolder
  }
  foreach ($r in @(Get-RevsSortedNewestFirst $revs)) {
    try {
      $null = Resolve-ProductionDocFolder $r.Path
      return $r.Path
    } catch {}
  }
  return (Get-LatestRevAmong2 $revs).Path
}

function Resolve-RevLetter([string]$targetDir, [object[]]$revRecords, [string]$DefaultLetter = '') {
  $revLetter = Get-RevLetterFromName (Split-Path -Leaf $targetDir)
  if (-not $revLetter -and $revRecords.Count -gt 0) {
    $latestRecRevLetter = $revRecords |
      ForEach-Object { Get-RevLetterFromName $_.Name } |
      Where-Object { $_ } |
      Select-Object -First 1
    if ($latestRecRevLetter) { $revLetter = $latestRecRevLetter }
  }
  if (-not $revLetter) {
    $pcbForRev = @(Get-AltiumPrjPcbFilesExcludingArchive $targetDir)
    foreach ($pf in $pcbForRev) {
      $r = Get-RevLetterFromName $pf.BaseName
      if ($r) { $revLetter = $r; break }
    }
  }
  if (-not $revLetter) {
    $default = ("$DefaultLetter").Trim()
    if ($default) {
      $revLetter = $default.ToUpperInvariant()
      Write-Warning (
        "Could not determine revision letter from folder name `"$(Split-Path -Leaf $targetDir)`" " +
        "or any .PrjPcb file under it (excluding paths under `"Archive`"); assuming Rev $revLetter."
      )
      return $revLetter
    }
    throw "Could not determine revision letter from folder name `"$(Split-Path -Leaf $targetDir)`" or any .PrjPcb file under it (excluding paths under `"Archive`")."
  }
  return $revLetter
}

function Resolve-PcbPartNumber([string]$canonical) {
  $segs = $canonical.Split('-')
  if ($segs.Count -ne 3) { return $canonical }
  $yy = switch ($segs[1]) {
    '10' { '20' }
    '15' { '25' }
    default { $segs[1] }
  }
  return ('{0}-{1}-{2}' -f $segs[0], $yy, $segs[2])
}

function GetItemMtimeUtc([string]$p) {
  try { return (Get-Item -LiteralPath $p).LastWriteTimeUtc } catch { return [datetime]::MinValue }
}

function Get-NewestPcbDocMtimeUtc([string]$folderPath) {
  # Prefer the newest *.PcbDoc under the folder (excluding Archive).
  # Falls back to folder mtime when no PCB is present.
  $pcbs = @(Get-AltiumPcbDocFilesExcludingArchive $folderPath)
  if ($pcbs.Count -gt 0) {
    $newest = $pcbs | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
    return $newest.LastWriteTimeUtc
  }
  return (GetItemMtimeUtc $folderPath)
}

function Select-BestPartFolder([string[]]$CandidatePaths) {
  $ranked = @(
    $CandidatePaths | ForEach-Object {
      [PSCustomObject]@{
        Path  = $_
        PcbMt = (Get-NewestPcbDocMtimeUtc $_)
        DirMt = (GetItemMtimeUtc $_)
      }
    } | Sort-Object @{ Expression = 'PcbMt'; Descending = $true }, @{ Expression = 'DirMt'; Descending = $true }
  )
  if ($ranked.Count -eq 0) { return $null }
  return $ranked[0].Path
}

function Get-LatestRevAmong2([object[]]$records) {
  # Wrap so a single sorted rev is still an array (.Count works); bare PSCustomObject does not.
  $sorted = @(Get-RevsSortedNewestFirst $records)
  if ($sorted.Count -gt 0) { return $sorted[0] }
  return $null
}

function Get-RevsSortedNewestFirst([object[]]$records) {
  # Sort rev records (objects with .Path, .Name, .Suffix) newest-first.
  # Numeric suffixes (e.g. "1", "2") win over text ones (e.g. "A", "B", "C1");
  # within each bucket, sort by suffix descending then by newest .PcbDoc mtime desc.
  if (-not $records -or $records.Count -eq 0) { return @() }
  $numeric = @()
  $text = @()
  foreach ($r in $records) {
    $s = $r.Suffix
    if ($s -match '^(\d+)$') {
      $numeric += [PSCustomObject]@{ Rec = $r; N = [decimal]$matches[1]; Mtime = (Get-NewestPcbDocMtimeUtc $r.Path) }
    }
    else {
      $text += [PSCustomObject]@{ Rec = $r; S = $s; Mtime = (Get-NewestPcbDocMtimeUtc $r.Path) }
    }
  }
  $sortedNumeric = @()
  if ($numeric.Count) {
    $sortedNumeric = @($numeric | Sort-Object @{ Expression = 'N'; Descending = $true }, @{ Expression = 'Mtime'; Descending = $true } | ForEach-Object { $_.Rec })
  }
  $sortedText = @()
  if ($text.Count) {
    $sortedText = @($text | Sort-Object @{ Expression = 'S'; Descending = $true }, @{ Expression = 'Mtime'; Descending = $true } | ForEach-Object { $_.Rec })
  }
  return @(@($sortedNumeric) + @($sortedText))
}

$parsed = Split-AltiumPartArg $PartNumber
$canonical = $parsed.Canonical
$tail = $parsed.NameTail

$searchRoots = @(Get-AltiumSearchRoots $Root)
$pppSegment = $canonical.Split('-')[0]

$candidates = New-Object System.Collections.Generic.List[string]

foreach ($searchRoot in $searchRoots) {
  # Directly under root
  Find-PartFoldersUnderParent -ParentDir $searchRoot -Canonical $canonical -OptionalNameTail $tail -OutCandidates $candidates

  # Under PPP <part> (e.g. ...\23. Altium\297\297-10-001)
  $pppJoin = Join-Path $searchRoot $pppSegment
  Find-PartFoldersUnderParent -ParentDir $pppJoin -Canonical $canonical -OptionalNameTail $tail -OutCandidates $candidates

  # Any other top-level folder that might wrap the part folder
  foreach ($top in Get-ChildItem -LiteralPath $searchRoot -Directory -ErrorAction SilentlyContinue) {
    if (Test-PartFolderTitleMatch $top.Name $canonical) { continue }

    Find-PartFoldersUnderParent -ParentDir $top.FullName -Canonical $canonical -OptionalNameTail $tail -OutCandidates $candidates
  }
}

if ($candidates.Count -eq 0) {
  $hintTail = if ($tail) {
    " or `"$canonical $tail`""
  }
  else { '' }
  $rootsLabel = $searchRoots -join '; '
  throw "No matching folder found for `"$canonical`"$hintTail under `"$rootsLabel`"."
}

$partDir = Select-BestPartFolder -CandidatePaths @($candidates)
if (-not $partDir) {
  throw "Failed to select a part folder among $($candidates.Count) candidate(s)."
}
if ($candidates.Count -gt 1) {
  Write-Host ("Multiple part folders matched; chose by newest .PcbDoc: {0}" -f (Split-Path -Leaf $partDir))
}

$subs = @(Get-ChildItem -LiteralPath $partDir -Directory -ErrorAction SilentlyContinue)
$revRecords = @()
foreach ($d in $subs) {
  $suf = Parse-RevSuffix $d.Name
  if ($null -ne $suf) {
    $revRecords += [PSCustomObject]@{ Path = $d.FullName; Suffix = $suf; Name = $d.Name }
  }
}

if ($revRecords.Count -eq 0) {
  $targetDir = $partDir
}
else {
  $latest = Get-LatestRevAmong2 @($revRecords)
  $targetDir = $latest.Path
}

Write-Host "Resolved design folder: $targetDir"

if ($Open -eq 'Explorer') {
  if ($revRecords.Count -gt 0) {
    Write-Host "Opening latest revision folder: $(Split-Path -Leaf $targetDir)"
  }
  else {
    Write-Host 'Opening (no Rev*/Ver* subfolders): part folder'
  }
  Start-ExplorerWindow -FolderPath $targetDir
  Write-Host "Done."
  exit 0
}

if ($Open -eq 'ProductionExplorer') {
  $prodWorkDir = Select-RevFolderWithProductionDoc -PartFolder $partDir -RevRecords $revRecords
  if ($prodWorkDir -ne $targetDir) {
    Write-Host "Production outputs taken from: $prodWorkDir"
  }
  $prodDocDir = Resolve-ProductionDocFolder $prodWorkDir
  Write-Host "Opening production document folder: $prodDocDir"
  Start-ExplorerWindow -FolderPath $prodDocDir
  Write-Host "Done."
  exit 0
}

if ($Open -eq 'PrjPcb' -or $Open -eq 'PcbDoc') {
  $isPcbDoc = ($Open -eq 'PcbDoc')
  $ext = if ($isPcbDoc) { '.PcbDoc' } else { '.PrjPcb' }
  $files = if ($isPcbDoc) {
    @(Get-AltiumPcbDocFilesExcludingArchive $targetDir)
  } else {
    @(Get-AltiumPrjPcbFilesExcludingArchive $targetDir)
  }
  if (-not $files -or $files.Count -eq 0) {
    throw "No $ext file found under `"$targetDir`" (paths under an `"Archive`" folder are ignored)."
  }
  $picked = $files | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
  Write-Host "Opening newest $ext`: $($picked.Name) (last write $($picked.LastWriteTime))"
  Start-Process -FilePath $picked.FullName
  # Altium owns the file association; try to bring its window forward
  # by matching the document filename (without extension) against the title bar.
  Show-AppWindowToFront -TitleHint $picked.BaseName
  Write-Host "Done."
  exit 0
}

if ($Open -eq 'ResolveJson') {
  # Emit resolved paths as JSON for other scripts to consume.
  #
  # The "primary" rev is the newest rev folder that actually contains a
  # production-document subfolder; if no rev has one, we fall back to the
  # plain newest rev. This matters because users sometimes start a new rev
  # folder (e.g. "Ver C1") with just firmware/notes, while the BOM and
  # fab outputs still live in the previous rev.
  $sortedRevs = @(Get-RevsSortedNewestFirst $revRecords)
  $candidates = @()
  if ($sortedRevs.Count -gt 0) {
    foreach ($r in $sortedRevs) {
      $pd = $null
      try { $pd = Resolve-ProductionDocFolder $r.Path } catch {}
      $candidates += [pscustomobject]@{
        name       = $r.Name
        path       = $r.Path
        suffix     = $r.Suffix
        prodDocDir = $pd
      }
    }
  } else {
    $pd = $null
    try { $pd = Resolve-ProductionDocFolder $partDir } catch {}
    $candidates += [pscustomobject]@{
      name       = (Split-Path -Leaf $partDir)
      path       = $partDir
      suffix     = $null
      prodDocDir = $pd
    }
  }
  $primary = ($candidates | Where-Object { $_.prodDocDir } | Select-Object -First 1)
  if (-not $primary) { $primary = $candidates[0] }

  # Re-derive revLetter from the picked rev's folder name (or fall back to a
  # PrjPcb scan under it). Avoid throwing if we can't determine it; downstream
  # callers can still use targetDir/prodDocDir.
  $primaryRevLetter = $null
  if ($primary.suffix) {
    $primaryRevLetter = $primary.suffix.ToUpperInvariant()
  } else {
    $primaryRevLetter = Get-RevLetterFromName $primary.name
  }
  if (-not $primaryRevLetter) {
    $pcbForRev = @(Get-AltiumPrjPcbFilesExcludingArchive $primary.path)
    foreach ($pf in $pcbForRev) {
      $r = Get-RevLetterFromName $pf.BaseName
      if ($r) { $primaryRevLetter = $r; break }
    }
  }

  $payload = [pscustomobject]@{
    canonical     = $canonical
    pcbPartNumber = (Resolve-PcbPartNumber $canonical)
    partDir       = $partDir
    targetDir     = $primary.path
    prodDocDir    = $primary.prodDocDir
    revLetter     = $primaryRevLetter
    revCandidates = $candidates
  } | ConvertTo-Json -Depth 6 -Compress
  [Console]::Out.WriteLine('---BEGIN-RESOLVE-JSON---')
  [Console]::Out.WriteLine($payload)
  [Console]::Out.WriteLine('---END-RESOLVE-JSON---')
  exit 0
}

# Fabrication / pick-place / check need outputs under a production folder; use
# the newest rev that has one (not necessarily the same as Explorer's $targetDir).
$prodWorkDir = $targetDir
if ($Open -in @('FabricationZip', 'GerberZip', 'PickPlaceZip', 'CheckFabrication')) {
  $prodWorkDir = Select-RevFolderWithProductionDoc -PartFolder $partDir -RevRecords $revRecords
  if ($prodWorkDir -ne $targetDir) {
    Write-Host "Production outputs taken from: $prodWorkDir"
  }
}

$prodDocDir = Resolve-ProductionDocFolder $prodWorkDir
Write-Host "Production document folder: $prodDocDir"

$revLetter = $null
if ($Open -ne 'PickPlaceZip' -or [string]::IsNullOrWhiteSpace($RevisionOverride)) {
  if ($Open -in @('FabricationZip', 'GerberZip')) {
    $revLetter = Resolve-RevLetter $prodWorkDir $revRecords -DefaultLetter 'A'
  }
  elseif ($Open -in @('PickPlaceZip', 'CheckFabrication')) {
    $revLetter = Resolve-RevLetter $prodWorkDir $revRecords
  }
}

if ($Open -eq 'PickPlaceZip') {
  $pickPlaceDir = Join-Path $prodDocDir 'Pick Place'
  if (-not (Test-Path -LiteralPath $pickPlaceDir -PathType Container)) {
    throw "Pick Place folder not found at `"$pickPlaceDir`"."
  }

  $ppFiles = @(Get-ChildItem -LiteralPath $pickPlaceDir -File -ErrorAction SilentlyContinue)
  $csvFiles = @($ppFiles | Where-Object { $_.Extension -ieq '.csv' })
  $txtFiles = @($ppFiles | Where-Object { $_.Extension -ieq '.txt' })

  if ($csvFiles.Count -gt 1) {
    $names = ($csvFiles | Sort-Object Name | ForEach-Object { $_.Name }) -join ', '
    throw "More than one .csv file in Pick Place folder: $names"
  }
  if ($txtFiles.Count -gt 1) {
    $names = ($txtFiles | Sort-Object Name | ForEach-Object { $_.Name }) -join ', '
    throw "More than one .txt file in Pick Place folder: $names"
  }

  $altiumRevForZip = $revLetter
  if (-not $altiumRevForZip -and -not [string]::IsNullOrWhiteSpace($RevisionOverride)) {
    try { $altiumRevForZip = Resolve-RevLetter $prodWorkDir $revRecords } catch { $altiumRevForZip = $null }
  }
  $zipRev = Resolve-PickPlaceZipRevisionLabel -AltiumRevision $altiumRevForZip -RevisionOverride $RevisionOverride
  if ($altiumRevForZip) {
    Write-Host "  Revision (Altium):          $altiumRevForZip"
  }
  if (-not [string]::IsNullOrWhiteSpace($RevisionOverride)) {
    Write-Host "  Revision (zip name):      $zipRev"
  }

  # If P&P source files are named only as a part number (e.g. 239-10-001.csv),
  # rename to "<PN> Pick and Place Rev <rev>" before zipping.
  $ppRenameTargets = @($csvFiles + $txtFiles | Where-Object {
      $_.BaseName -match '^\d{3}-\d{2}-\d{3}$'
    })
  foreach ($f in $ppRenameTargets) {
    $pnForName = $f.BaseName
    $newName = '{0} Pick and Place Rev {1}{2}' -f $pnForName, $zipRev, $f.Extension
    $dest = Join-Path $f.DirectoryName $newName
    if (
      (Test-Path -LiteralPath $dest) -and
      -not [string]::Equals($f.FullName, $dest, [StringComparison]::OrdinalIgnoreCase)
    ) {
      throw (
        "Cannot rename Pick Place file from '$($f.Name)' to '$newName': " +
        "destination already exists at `"$dest`"."
      )
    }
    Write-Host ("  Renaming Pick Place file: '{0}' -> '{1}'" -f $f.Name, $newName)
    Rename-Item -LiteralPath $f.FullName -NewName $newName -Force
  }

  $zipBase = "Pick Place $canonical Rev $zipRev"
  $zipPath = Join-Path $prodDocDir ($zipBase + '.zip')
  Write-Host "Zip target: $zipPath"

  Invoke-ZipOverwritePromptIfNeeded -ZipPath $zipPath -ForceZipOverwrite:$ForceZipOverwrite

  New-AltiumFolderZip -SourceDirectory $pickPlaceDir -DestinationZipPath $zipPath

  Write-Host "Created: $zipPath"
  Start-ExplorerWindow -SelectPath $zipPath
  exit 0
}

if ($Open -eq 'CheckFabrication') {
  $checkResult = Invoke-FabricationFolderChecks -ProdDocDir $prodDocDir `
    -TimestampSpreadMinutes $TimestampSpreadMinutes
  Show-FabricationFolderCheckReport -ProdDocDir $prodDocDir -RevLetter $revLetter `
    -TimestampSpreadMinutes $TimestampSpreadMinutes -CheckResult $checkResult
  $errs = @($checkResult.Errors)
  $warns = @($checkResult.Warnings)
  if ($errs.Count -eq 0 -and $warns.Count -eq 0) {
    exit 0
  }
  if ($errs.Count -gt 0) { exit 1 } else { exit 2 }
}

if ($Open -eq 'GerberZip') {
  $pcbPart = Resolve-PcbPartNumber $canonical
  $revLetterOnly = ($revLetter -replace '\d+$', '')
  if ([string]::IsNullOrWhiteSpace($revLetterOnly)) { $revLetterOnly = $revLetter }
  $zipBase = "Gerbers $pcbPart Rev $revLetterOnly"
  $zipPath = Join-Path $prodDocDir ($zipBase + '.zip')

  $needed = 'Gerber', 'NC Drill'
  $present = New-Object System.Collections.Generic.List[string]
  foreach ($n in $needed) {
    $p = Join-Path $prodDocDir $n
    if (Test-Path -LiteralPath $p -PathType Container) {
      $null = $present.Add($p)
    }
    else {
      Write-Warning "Missing folder: ``$n`` (expected at ``$p``). Continuing without it."
    }
  }

  if ($present.Count -eq 0) {
    throw "None of the required folders (Gerber, NC Drill) were found under `"$prodDocDir`"."
  }

  $checkResult = Invoke-FabricationFolderChecks -ProdDocDir $prodDocDir `
    -TimestampSpreadMinutes $TimestampSpreadMinutes -IncludePfsChecks:$false -AllowTwoNcDrillFiles
  Show-FabricationFolderCheckReport -ProdDocDir $prodDocDir -RevLetter $revLetter `
    -TimestampSpreadMinutes $TimestampSpreadMinutes -CheckResult $checkResult `
    -ReportTitle 'Gerber folder check'
  if ($checkResult.Errors.Count -gt 0) {
    $detail = ($checkResult.Errors | ForEach-Object { $_ }) -join ' '
    throw "Gerber zip aborted: $($checkResult.Errors.Count) check(s) failed. $detail"
  }

  Write-Host ''
  Write-Host "Zip target: $zipPath"

  Invoke-ZipOverwritePromptIfNeeded -ZipPath $zipPath -ForceZipOverwrite:$ForceZipOverwrite
  if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
  }

  Compress-Archive -Path $present -DestinationPath $zipPath -CompressionLevel Optimal -Force

  Write-Host "Created: $zipPath"
  Write-Host "Included: $($present | ForEach-Object { Split-Path -Leaf $_ })"
  Start-ExplorerWindow -SelectPath $zipPath
  exit 0
}

# Open = FabricationZip
if ($Open -ne 'FabricationZip') {
  throw "Unsupported -Open value for zip output: $Open"
}
$pcbPart = Resolve-PcbPartNumber $canonical
# Fabrication zip names use the letter portion only (e.g. "Rev C"), even when
# the source folder is "Ver C1" — the bare board's rev letter doesn't
# include the assembly's trailing digit. Pick & Place keeps the full rev.
$revLetterOnly = ($revLetter -replace '\d+$', '')
if ([string]::IsNullOrWhiteSpace($revLetterOnly)) { $revLetterOnly = $revLetter }
$zipBase = "Fabrication $pcbPart Rev $revLetterOnly"
$zipPath = Join-Path $prodDocDir ($zipBase + '.zip')

$needed = 'Gerber', 'NC Drill', 'PFS'
$present = New-Object System.Collections.Generic.List[string]
foreach ($n in $needed) {
  $p = Join-Path $prodDocDir $n
  if (Test-Path -LiteralPath $p -PathType Container) {
    $null = $present.Add($p)
  }
  else {
    Write-Warning "Missing folder: ``$n`` (expected at ``$p``). Continuing without it."
  }
}

if ($present.Count -eq 0) {
  throw "None of the required folders (Gerber, NC Drill, PFS) were found under `"$prodDocDir`"."
}

$checkResult = Invoke-FabricationFolderChecks -ProdDocDir $prodDocDir `
  -TimestampSpreadMinutes $TimestampSpreadMinutes -AllowTwoNcDrillFiles
Show-FabricationFolderCheckReport -ProdDocDir $prodDocDir -RevLetter $revLetter `
  -TimestampSpreadMinutes $TimestampSpreadMinutes -CheckResult $checkResult
if ($checkResult.Errors.Count -gt 0) {
  $detail = ($checkResult.Errors | ForEach-Object { $_ }) -join ' '
  throw "Fabrication zip aborted: $($checkResult.Errors.Count) check(s) failed. $detail"
}

Write-Host ''
Write-Host "Zip target: $zipPath"

Invoke-ZipOverwritePromptIfNeeded -ZipPath $zipPath -ForceZipOverwrite:$ForceZipOverwrite
if (Test-Path -LiteralPath $zipPath) {
  Remove-Item -LiteralPath $zipPath -Force
}

Compress-Archive -Path $present -DestinationPath $zipPath -CompressionLevel Optimal -Force

Write-Host "Created: $zipPath"
Write-Host "Included: $($present | ForEach-Object { Split-Path -Leaf $_ })"
Start-ExplorerWindow -SelectPath $zipPath
