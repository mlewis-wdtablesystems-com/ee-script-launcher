<#
.SYNOPSIS
  Correct Pick & Place CSV and TXT in the production Pick Place folder from Arena BOM + PJM library comments.

.DESCRIPTION
  Given a PCBA / assembly part number:
    1. Resolves the Altium production document folder and locates the single .csv and .txt
       in the "Pick Place" subfolder (same rules as Open-AltiumPartFolder PickPlaceZip).
    2. Fetches the Arena assembly BOM for the chosen revision (working or most recent released).
    3. For each BOM line, maps reference designators to Arena item description (description
       when set, else item name) and part number.
    4. Updates both pick-place files:
         - Description := Arena description (when a designator matches a BOM line).
         - Comment := PJM_Component_Database_Lib.xls Comment for 01-PJM Part_Number when found;
           else keep existing comment if it reasonably matches the Arena description;
           else Comment := Arena description.
         - Skips lines whose designator/comment/description/footprint indicate Fiducial.
         - Removes lines whose designators belong to Arena BOM lines marked Not Fitted / N/F / DNF.

  Original files are copied to the production folder subfolder
  "Pick Place Archive - DO NOT USE" (*.bak) before overwrite.

.PARAMETER PartNumber
  Assembly part number used for Arena BOM lookup and Altium folder resolution (e.g. 272-15-001).

.PARAMETER RevisionMode
  working = BOM on the working revision (status 0).
  effective_latest = BOM on the newest effective (released) revision (status 1).

.PARAMETER PjmDatabasePath
  Path to PJM_Component_Database_Lib.xls (default: standard PJM library location).

.EXAMPLE
  .\Correct-PickPlaceFromArenaBom.ps1 -PartNumber 272-15-001
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [ValidateSet('working', 'effective_latest')] [string]$RevisionMode = 'working',
  [Parameter(Mandatory = $false)]
  [string]$PjmDatabasePath = (
    'C:\Users\mlewis\OneDrive - Walker Digital\WTDS-RDAustralia - Documents\23. Altium\Altium Library\PJM\PJM_Component_Database_Lib.xls'
  )
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$openScript = Join-Path $scriptsDir 'Open-AltiumPartFolder.ps1'
$checksLib  = Join-Path $scriptsDir 'Bom-ArenaAltiumDescriptionChecks.ps1'
foreach ($p in @($openScript, $checksLib)) {
  if (-not (Test-Path -LiteralPath $p)) { throw "Required script or library missing: $p" }
}
. $checksLib

if ([string]::IsNullOrWhiteSpace($env:ARENA_EMAIL)) { throw 'Set ARENA_EMAIL.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_PASSWORD)) { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_WORKSPACE_ID)) {
  throw 'Set ARENA_WORKSPACE_ID (launcher workspace dropdown sets this per run).'
}
if (-not (Test-Path -LiteralPath $PjmDatabasePath)) {
  throw "PJM component database not found: $PjmDatabasePath"
}

function Invoke-ArenaJson {
  param(
    [Parameter(Mandatory = $true)] [string]$Uri,
    [string]$Method = 'Get',
    [hashtable]$Headers,
    [string]$Body,
    [string]$ContentType = 'application/json'
  )
  $params = @{ Uri = $Uri; Method = $Method; UseBasicParsing = $true }
  if ($Headers) { $params.Headers = $Headers }
  if ($Body) { $params.Body = $Body; $params.ContentType = $ContentType }
  $resp = Invoke-WebRequest @params
  $bytes = $resp.RawContentStream.ToArray()
  $text = [System.Text.Encoding]::UTF8.GetString($bytes)
  if ([string]::IsNullOrWhiteSpace($text)) { return $null }
  return $text | ConvertFrom-Json
}

function Select-ArenaAssemblyRevision {
  param(
    [Parameter(Mandatory = $true)] [object[]]$RevResults,
    [Parameter(Mandatory = $true)] [string]$ModeKey
  )
  if ($ModeKey -eq 'working') {
    $w = $RevResults | Where-Object { $_.status -eq 0 } | Select-Object -First 1
    if (-not $w) { throw 'No working revision (status 0) on this Arena item.' }
    return $w
  }
  $eff = @($RevResults | Where-Object { $_.status -eq 1 })
  if ($eff.Count -eq 0) {
    throw 'No effective (released) revision (status 1). Choose working revision or release the item first.'
  }
  return $eff | Sort-Object { $_.effectiveDateTime } -Descending | Select-Object -First 1
}

function Split-RefDesignators([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return @() }
  return @(($s -split '[,;]') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Test-ArenaBomLineIsNotFitted {
  param(
    [string]$PartNumber,
    [string]$Name,
    [string]$Notes
  )
  $blob = ("$PartNumber $Name $Notes").Trim()
  if (-not $blob) { return $false }
  return [bool]($blob -match '(?i)\b(not\s*fitted|n\s*/\s*f|dnf|do\s+not\s+fit|not\s+installed)\b')
}

function Test-PickPlaceRowIsFiducial {
  param([hashtable]$Row)
  foreach ($k in @('Designator', 'Comment', 'Description', 'Footprint')) {
    $v = $Row[$k]
    if ($v -and ("$v" -match '(?i)fiducial')) { return $true }
  }
  return $false
}

function Get-PickPlaceCommentForPart {
  param(
    [string]$ArenaPartNumber,
    [string]$ArenaDescription,
    [string]$ExistingComment,
    [hashtable]$PjmComments
  )
  $pn = $ArenaPartNumber.Trim()
  if ($pn -and $PjmComments.ContainsKey($pn)) {
    $c = ("$($PjmComments[$pn])").Trim()
    if ($c) { return $c }
  }
  $existing = ("$ExistingComment").Trim()
  $arena = ("$ArenaDescription").Trim()
  if ($existing -and $arena) {
    $na = Normalize-BomDescriptionText $existing
    $nb = Normalize-BomDescriptionText $arena
    if ($na -eq $nb) { return $existing }
    if ($na.Contains($nb) -or $nb.Contains($na)) { return $existing }
    if (Test-BomLooseTextOverlap -na $na -nb $nb) { return $existing }
  }
  return $arena
}

function Get-PjmPartNumberCommentMap {
  param([string]$XlsPath)
  $map = @{}
  $excel = $null
  $wb = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($XlsPath, 0, $true)
    foreach ($ws in @($wb.Worksheets)) {
      $used = $ws.UsedRange
      if (-not $used) { continue }
      $rowCount = [int]$used.Rows.Count
      $colCount = [int]$used.Columns.Count
      $values = $used.Value2
      $headerRow = -1
      $colPn = -1
      $colComment = -1
      $scan = [Math]::Min($rowCount, 30)
      for ($r = 1; $r -le $scan; $r++) {
        for ($c = 1; $c -le $colCount; $c++) {
          $cell = ("" + $values[$r, $c]).Trim()
          if ($cell -eq '01-PJM Part_Number' -or $cell -match '(?i)^01-PJM\s*Part[_ ]?Number$') {
            $headerRow = $r
            $colPn = $c
            break
          }
        }
        if ($headerRow -ge 0) { break }
      }
      if ($headerRow -lt 0) { continue }
      for ($c = 1; $c -le $colCount; $c++) {
        $h = ("" + $values[$headerRow, $c]).Trim()
        if ($h -imatch '^Comment$') { $colComment = $c; break }
      }
      if ($colComment -lt 0) { continue }
      for ($r = $headerRow + 1; $r -le $rowCount; $r++) {
        $pn = ("" + $values[$r, $colPn]).Trim()
        if (-not $pn) { continue }
        $cmt = ("" + $values[$r, $colComment]).Trim()
        if (-not $map.ContainsKey($pn)) {
          $map[$pn] = $cmt
        }
      }
    }
  }
  finally {
    if ($wb) {
      try { $wb.Close($false) } catch {}
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
    if ($excel) {
      try { $excel.Quit() } catch {}
      [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
  return $map
}

function Find-PickPlaceHeaderLineIndex {
  param([string[]]$Lines)
  for ($i = 0; $i -lt $Lines.Count; $i++) {
    $l = $Lines[$i]
    if ($l -match '(?i)Designator' -and $l -match '(?i)Comment') {
      return $i
    }
    if ($l -match '(?i)^"Designator"' -and $l -match '(?i)Comment') {
      return $i
    }
  }
  return -1
}

function ConvertFrom-PickPlaceCsvRows {
  param(
    [string]$HeaderLine,
    [string[]]$DataLines
  )
  $block = @($HeaderLine) + @($DataLines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
  $text = $block -join "`r`n"
  return @($text | ConvertFrom-Csv)
}

function ConvertTo-PickPlaceCsvLine {
  param([psobject]$Row)
  $cols = @($Row.PSObject.Properties | ForEach-Object { $_.Name })
  $parts = foreach ($c in $cols) {
    $v = $Row.$c
    if ($null -eq $v) { $v = '' }
    $s = "$v"
    if ($s -match '[",\r\n]') {
      '"' + ($s -replace '"', '""') + '"'
    }
    else {
      $s
    }
  }
  return ($parts -join ',')
}

function Get-TxtColumnLayout {
  param([string]$HeaderLine)
  $known = @(
    'Designator', 'Comment', 'Layer', 'Footprint',
    'Center-X(mil)', 'Center-Y(mil)', 'Center-X(mm)', 'Center-Y(mm)',
    'Rotation', 'Description', '01-PJM Part_Number', '03-Description General'
  )
  $present = New-Object System.Collections.Generic.List[object]
  foreach ($name in $known) {
    $idx = $HeaderLine.IndexOf($name, [System.StringComparison]::OrdinalIgnoreCase)
    if ($idx -ge 0) {
      $present.Add([pscustomobject]@{ Name = $name; Start = $idx })
    }
  }
  if ($present.Count -eq 0) { throw 'Could not parse TXT column layout from header.' }
  $sorted = @($present | Sort-Object Start)
  $layout = @()
  for ($i = 0; $i -lt $sorted.Count; $i++) {
    $start = $sorted[$i].Start
    $end = if ($i -lt $sorted.Count - 1) { $sorted[$i + 1].Start } else { $HeaderLine.Length }
    $width = [Math]::Max(1, $end - $start)
    $layout += [pscustomobject]@{ Name = $sorted[$i].Name; Start = $start; Width = $width }
  }
  return $layout
}

function Read-TxtPickPlaceRow {
  param(
    [string]$Line,
    [array]$Layout
  )
  $row = @{}
  foreach ($col in $Layout) {
    $slice = if ($col.Start -lt $Line.Length) {
      $len = [Math]::Min($col.Width, $Line.Length - $col.Start)
      $Line.Substring($col.Start, $len)
    } else { '' }
    $row[$col.Name] = $slice.Trim().Trim('"')
  }
  return $row
}

function Format-TxtPickPlaceRow {
  param(
    [hashtable]$Row,
    [array]$Layout,
    [int]$LineLength
  )
  $chars = [char[]]::new([Math]::Max($LineLength, 80))
  for ($i = 0; $i -lt $chars.Length; $i++) { $chars[$i] = ' ' }
  foreach ($col in $Layout) {
    $val = ("$($Row[$col.Name])").Trim()
    if ($col.Name -eq 'Description' -and $val -match '\s') {
      $val = '"' + $val + '"'
    }
    $start = $col.Start
    $maxLen = [Math]::Min($col.Width, $chars.Length - $start)
    if ($maxLen -lt 1) { continue }
    $chunk = if ($val.Length -gt $maxLen) { $val.Substring(0, $maxLen) } else { $val }
    for ($j = 0; $j -lt $chunk.Length; $j++) {
      if ($start + $j -lt $chars.Length) { $chars[$start + $j] = $chunk[$j] }
    }
  }
  return -join $chars
}

$PickPlaceArchiveFolderName = 'Pick Place Archive - DO NOT USE'

function Backup-PickPlaceFileBeforeUpdate {
  param(
    [Parameter(Mandatory = $true)] [string]$SourcePath,
    [Parameter(Mandatory = $true)] [string]$ArchiveDir
  )
  if (-not (Test-Path -LiteralPath $ArchiveDir)) {
    New-Item -ItemType Directory -Path $ArchiveDir -Force | Out-Null
  }
  $bakName = (Split-Path -Leaf $SourcePath) + '.bak'
  $bakPath = Join-Path $ArchiveDir $bakName
  Copy-Item -LiteralPath $SourcePath -Destination $bakPath -Force
  return $bakPath
}

function Update-PickPlaceTextFile {
  param(
    [string]$Path,
    [string]$ArchiveDir,
    [hashtable]$DesignatorMap,
    [hashtable]$RemoveDesignators,
    [hashtable]$PjmComments
  )
  $stats = @{
    Updated         = 0
    Removed         = 0
    FiducialSkipped = 0
    NoArenaMatch    = 0
  }
  $utf8 = New-Object System.Text.UTF8Encoding $false
  $lines = [System.IO.File]::ReadAllLines($Path, $utf8)
  $headerIdx = Find-PickPlaceHeaderLineIndex -Lines $lines
  if ($headerIdx -lt 0) { throw "No Designator/Comment header row in: $Path" }

  $isCsv = $Path -match '\.csv$'
  $headerLine = $lines[$headerIdx]
  $dataLines = @()
  for ($i = $headerIdx + 1; $i -lt $lines.Count; $i++) {
    $l = $lines[$i]
    if ([string]::IsNullOrWhiteSpace($l)) { continue }
    if ($l.StartsWith('=')) { break }
    $dataLines += $l
  }

  if ($isCsv) {
    $parsed = ConvertFrom-PickPlaceCsvRows -HeaderLine $headerLine -DataLines $dataLines
    $outData = New-Object System.Collections.Generic.List[string]
    $needDescCol = $headerLine -notmatch '(?i)Description'
    foreach ($row in $parsed) {
      $des = ("$($row.Designator)").Trim()
      if (-not $des) { continue }
      if ($RemoveDesignators.ContainsKey($des)) {
        $stats.Removed++
        continue
      }
      $ht = @{}
      foreach ($p in $row.PSObject.Properties) { $ht[$p.Name] = $p.Value }
      if (Test-PickPlaceRowIsFiducial -Row $ht) {
        $stats.FiducialSkipped++
        $outData.Add((ConvertTo-PickPlaceCsvLine -Row $row))
        continue
      }
      if (-not $DesignatorMap.ContainsKey($des)) {
        $stats.NoArenaMatch++
        $outData.Add((ConvertTo-PickPlaceCsvLine -Row $row))
        continue
      }
      $arena = $DesignatorMap[$des]
      $existingComment = if ($ht.ContainsKey('Comment')) { $ht['Comment'] } else { '' }
      $newComment = Get-PickPlaceCommentForPart `
        -ArenaPartNumber $arena.PartNumber `
        -ArenaDescription $arena.Description `
        -ExistingComment $existingComment `
        -PjmComments $PjmComments
      $row.Comment = $newComment
      if ($row.PSObject.Properties.Name -contains 'Description') {
        $row.Description = $arena.Description
      }
      else {
        $row | Add-Member -NotePropertyName Description -NotePropertyValue $arena.Description -Force
        $needDescCol = $true
      }
      $stats.Updated++
      $outData.Add((ConvertTo-PickPlaceCsvLine -Row $row))
    }
    if ($needDescCol -and $headerLine -notmatch '(?i)Description') {
      $headerLine = $headerLine.TrimEnd()
      if ($headerLine.EndsWith('"')) {
        $headerLine += ',"Description"'
      }
      else {
        $headerLine += ',Description'
      }
    }
    $newLines = @()
    if ($headerIdx -gt 0) { $newLines += $lines[0..($headerIdx - 1)] }
    $newLines += $headerLine
    $newLines += $outData
    $bak = Backup-PickPlaceFileBeforeUpdate -SourcePath $Path -ArchiveDir $ArchiveDir
    [System.IO.File]::WriteAllLines($Path, $newLines, $utf8)
    Write-Host "  Wrote CSV (backup: $bak)"
    return $stats
  }

  $layout = Get-TxtColumnLayout -HeaderLine $headerLine
  $sampleLen = if ($dataLines.Count -gt 0) { ($dataLines | Measure-Object -Maximum Length).Maximum } else { $headerLine.Length }
  $outTxt = New-Object System.Collections.Generic.List[string]
  foreach ($line in $dataLines) {
    $row = Read-TxtPickPlaceRow -Line $line -Layout $layout
    $des = ("$($row['Designator'])").Trim()
    if (-not $des) { continue }
    if ($RemoveDesignators.ContainsKey($des)) {
      $stats.Removed++
      continue
    }
    if (Test-PickPlaceRowIsFiducial -Row $row) {
      $stats.FiducialSkipped++
      $outTxt.Add((Format-TxtPickPlaceRow -Row $row -Layout $layout -LineLength $sampleLen))
      continue
    }
    if (-not $DesignatorMap.ContainsKey($des)) {
      $stats.NoArenaMatch++
      $outTxt.Add((Format-TxtPickPlaceRow -Row $row -Layout $layout -LineLength $sampleLen))
      continue
    }
    $arena = $DesignatorMap[$des]
    $existingComment = if ($row.ContainsKey('Comment')) { $row['Comment'] } else { '' }
    $row['Comment'] = Get-PickPlaceCommentForPart `
      -ArenaPartNumber $arena.PartNumber `
      -ArenaDescription $arena.Description `
      -ExistingComment $existingComment `
      -PjmComments $PjmComments
    if ($row.ContainsKey('Description')) {
      $row['Description'] = $arena.Description
    }
    $stats.Updated++
    $outTxt.Add((Format-TxtPickPlaceRow -Row $row -Layout $layout -LineLength $sampleLen))
  }
  $newLines = @()
  if ($headerIdx -gt 0) { $newLines += $lines[0..($headerIdx - 1)] }
  $newLines += $headerLine
  $newLines += $outTxt
  $bak = Backup-PickPlaceFileBeforeUpdate -SourcePath $Path -ArchiveDir $ArchiveDir
  [System.IO.File]::WriteAllLines($Path, $newLines, $utf8)
  Write-Host "  Wrote TXT (backup: $bak)"
  return $stats
}

# ---------- Resolve Altium Pick Place folder ----------

Write-Host "=== Resolving Altium folder for $PartNumber ==="
$resolveOut = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File $openScript -PartNumber $PartNumber.Trim() -Open ResolveJson 2>&1
if ($LASTEXITCODE -ne 0) {
  throw "Open-AltiumPartFolder.ps1 -Open ResolveJson failed (exit $LASTEXITCODE):`n$($resolveOut -join [Environment]::NewLine)"
}
$joined = ($resolveOut | ForEach-Object { "$_" }) -join "`n"
$rxJson = [regex]'---BEGIN-RESOLVE-JSON---\s*(?<j>\{.*\})\s*---END-RESOLVE-JSON---'
$jm = $rxJson.Match($joined)
if (-not $jm.Success) { throw "Could not find resolve JSON markers.`n$joined" }
$resolved = $jm.Groups['j'].Value | ConvertFrom-Json
$prodDoc = $resolved.prodDocDir
Write-Host ("  Production folder : {0}" -f $prodDoc)

$pickPlaceDir = Join-Path $prodDoc 'Pick Place'
if (-not (Test-Path -LiteralPath $pickPlaceDir -PathType Container)) {
  throw "Pick Place folder not found at `"$pickPlaceDir`"."
}
$ppFiles = @(Get-ChildItem -LiteralPath $pickPlaceDir -File -ErrorAction SilentlyContinue)
$csvFiles = @($ppFiles | Where-Object { $_.Extension -ieq '.csv' })
$txtFiles = @($ppFiles | Where-Object { $_.Extension -ieq '.txt' })
if ($csvFiles.Count -ne 1) {
  throw "Expected exactly one .csv in Pick Place folder; found $($csvFiles.Count)."
}
if ($txtFiles.Count -ne 1) {
  throw "Expected exactly one .txt in Pick Place folder; found $($txtFiles.Count)."
}
$csvPath = $csvFiles[0].FullName
$txtPath = $txtFiles[0].FullName
$pickPlaceArchiveDir = Join-Path $prodDoc $PickPlaceArchiveFolderName
Write-Host ("  CSV: {0}" -f $csvPath)
Write-Host ("  TXT: {0}" -f $txtPath)
Write-Host ("  Backups: {0}" -f $pickPlaceArchiveDir)

# ---------- Arena BOM + item descriptions ----------

Write-Host ''
Write-Host '=== Arena login + BOM ==='
$loginBody = @{
  email       = $env:ARENA_EMAIL
  password    = $env:ARENA_PASSWORD
  workspaceId = [long]$env:ARENA_WORKSPACE_ID
} | ConvertTo-Json
$login = Invoke-ArenaJson -Uri 'https://api.arenasolutions.com/v1/login' -Method Post -Body $loginBody
$sessionId = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sessionId)) { $sessionId = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sessionId)) { throw 'Login response did not include arenaSessionId.' }
$arenaHeaders = @{
  arena_session_id = $sessionId
  'Content-Type'   = 'application/json'
}

$pnEnc = [Uri]::EscapeDataString($PartNumber.Trim())
$lookup = Invoke-ArenaJson -Uri "https://api.arenasolutions.com/v1/items?number=$pnEnc&limit=1" -Headers $arenaHeaders
if (-not $lookup.results -or $lookup.results.Count -eq 0) {
  throw "No Arena item found for part number '$($PartNumber.Trim())'."
}
$parentItem = $lookup.results[0]
$parentGuid = $parentItem.guid
$revs = Invoke-ArenaJson -Uri "https://api.arenasolutions.com/v1/items/$parentGuid/revisions" -Headers $arenaHeaders
if (-not $revs.results -or $revs.results.Count -eq 0) {
  throw "No revisions returned for Arena item '$($PartNumber.Trim())'."
}
$assemblyRev = Select-ArenaAssemblyRevision -RevResults @($revs.results) -ModeKey $RevisionMode
$assemblyRevGuid = ("$($assemblyRev.guid)").Trim()
$assemblyRevNum = if ($null -ne $assemblyRev.number) { ("$($assemblyRev.number)").Trim() } else { '' }
$revLabel = if ($RevisionMode -eq 'working') { 'working' } else { 'most recent released' }
Write-Host ("  Item: {0}  ({1})" -f $parentItem.number, $parentItem.name)
Write-Host ("  BOM revision ({0}): {1}  guid={2}" -f $revLabel, $(if ($assemblyRevNum) { $assemblyRevNum } else { '(no number)' }), $assemblyRevGuid)

$bom = Invoke-ArenaJson -Uri "https://api.arenasolutions.com/v1/items/$assemblyRevGuid/bom" -Headers $arenaHeaders
if (-not $bom.results) {
  throw "Arena BOM returned no lines for $($PartNumber.Trim()) on $revLabel revision."
}

$itemDescCache = @{}
function Get-ArenaItemDescriptionText {
  param([string]$ItemNumber)
  $key = $ItemNumber.Trim()
  if (-not $key) { return '' }
  if ($itemDescCache.ContainsKey($key)) { return $itemDescCache[$key] }
  $enc = [Uri]::EscapeDataString($key)
  $viewEnc = [Uri]::EscapeDataString('full')
  $items = Invoke-ArenaJson -Uri "https://api.arenasolutions.com/v1/items?number=$enc&limit=1&responseview=$viewEnc" -Headers $arenaHeaders
  $side = ''
  if ($items.results -and $items.results.Count -gt 0) {
    $it = $items.results[0]
    $desc = if ($null -ne $it.description) { ("$($it.description)").Trim() } else { '' }
    $nm = if ($null -ne $it.name) { ("$($it.name)").Trim() } else { '' }
    $side = if ($desc) { $desc } else { $nm }
  }
  $itemDescCache[$key] = $side
  return $side
}

$designatorMap = @{}
$removeDesignators = @{}
$nfLines = 0
foreach ($r in $bom.results) {
  $pn = ("$($r.item.number)").Trim()
  $name = ("$($r.item.name)").Trim()
  $notes = if ($null -ne $r.notes) { ("$($r.notes)").Trim() } else { '' }
  $isNf = Test-ArenaBomLineIsNotFitted -PartNumber $pn -Name $name -Notes $notes
  if ($isNf) { $nfLines++ }
  $desc = Get-ArenaItemDescriptionText -ItemNumber $pn
  if (-not $desc) { $desc = $name }
  foreach ($des in @(Split-RefDesignators ("$($r.refDes)"))) {
    if (-not $des) { continue }
    if ($isNf) {
      $removeDesignators[$des] = $true
      continue
    }
    $designatorMap[$des] = [pscustomobject]@{
      PartNumber  = $pn
      Description = $desc
    }
  }
}
Write-Host ("  BOM lines: {0}  designators mapped: {1}  NF/DNF designators to remove: {2}" -f `
  $bom.results.Count, $designatorMap.Count, $removeDesignators.Count)

Write-Host ''
Write-Host '=== Loading PJM component database comments ==='
$pjmComments = Get-PjmPartNumberCommentMap -XlsPath $PjmDatabasePath
Write-Host ("  PJM part numbers with Comment: {0}" -f $pjmComments.Count)

# ---------- Update both files ----------

Write-Host ''
Write-Host '=== Updating Pick Place files ==='
$csvStats = Update-PickPlaceTextFile -Path $csvPath -ArchiveDir $pickPlaceArchiveDir `
  -DesignatorMap $designatorMap -RemoveDesignators $removeDesignators -PjmComments $pjmComments
$txtStats = Update-PickPlaceTextFile -Path $txtPath -ArchiveDir $pickPlaceArchiveDir `
  -DesignatorMap $designatorMap -RemoveDesignators $removeDesignators -PjmComments $pjmComments

Write-Host ''
Write-Host '=== Summary ===' -ForegroundColor Green
foreach ($label in @('CSV', 'TXT')) {
  $s = if ($label -eq 'CSV') { $csvStats } else { $txtStats }
  Write-Host "  $label :"
  Write-Host ("    Updated (Comment/Description): {0}" -f $s.Updated)
  Write-Host ("    Removed (Arena NF/DNF):        {0}" -f $s.Removed)
  Write-Host ("    Fiducial unchanged:            {0}" -f $s.FiducialSkipped)
  Write-Host ("    No Arena match (kept):         {0}" -f $s.NoArenaMatch)
}
Write-Host ''
Write-Host 'Done.' -ForegroundColor Green
exit 0
