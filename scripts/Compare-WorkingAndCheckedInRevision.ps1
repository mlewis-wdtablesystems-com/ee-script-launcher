<#
.SYNOPSIS
  Compare a PCBA and bare-board item's working revision against the most recent checked-in revision.

.DESCRIPTION
  For a PCBA part number (PPP-10-XXX / PPP-15-XXX) and its mapped bare board
  (PPP-20-XXX / PPP-25-XXX), compares working (status 0) vs newest effective /
  checked-in (status 1) revision and reports:

    - Revision number differences
    - File association differences (by category; title / edition / file number)
    - BOM line differences (part number, quantity, reference designators, notes)
    - Revision-level item fields (name, description, lifecycle phase, UOM)
      assemblyType differences are ignored. A blank working revision number is
      ignored when files, BOM, and other fields match checked-in.

  Read-only. Exit 0 when working matches checked-in; exit 2 when any difference is found.

.PARAMETER PartNumber
  PCBA internal part number, e.g. 251-10-001 or 251-15-001.

.PARAMETER DeferSummary
  When set (used by Check-ArenaPcbaFiles.ps1), emit the summary between
  ---BEGIN-COMPARE-SUMMARY--- / ---END-COMPARE-SUMMARY--- markers instead of
  printing --- Summary --- inline. Exit code is unchanged.

.EXAMPLE
  .\Compare-WorkingAndCheckedInRevision.ps1 -PartNumber 251-10-001
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [switch]$DeferSummary,
  [string]$Email = $env:ARENA_EMAIL,
  [string]$WorkspaceId = $env:ARENA_WORKSPACE_ID,
  [string]$BaseUrl = 'https://api.arenasolutions.com/v1'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $scriptsDir 'AltiumRevisionHelpers.ps1')

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

$password = $env:ARENA_PASSWORD
if ([string]::IsNullOrWhiteSpace($Email))       { throw 'Set ARENA_EMAIL or pass -Email.' }
if ([string]::IsNullOrWhiteSpace($password))    { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($WorkspaceId)) { throw 'Set ARENA_WORKSPACE_ID or pass -WorkspaceId.' }

$parts = ConvertTo-PcbaAndBareBoardPartNumbers -PartNumber $PartNumber
$pcbaNumber = $parts.Pcba
$bareNumber = $parts.Bare

$apiRoot = $BaseUrl.TrimEnd('/')

$loginBody = @{ email = $Email; password = $password; workspaceId = [long]$WorkspaceId } | ConvertTo-Json
try {
  $login = Invoke-ArenaJson -Uri "$apiRoot/login" -Method Post -Body $loginBody
} catch { throw "Arena login failed: $($_.Exception.Message)" }

$sid = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sid)) { $sid = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sid)) { throw 'Login response did not include arenaSessionId.' }

$headers = @{ arena_session_id = $sid; 'Content-Type' = 'application/json' }

$differences = New-Object System.Collections.Generic.List[string]

function Get-WorkingAndCheckedInRevisions {
  param([string]$ItemNumber)
  $pnEnc = [Uri]::EscapeDataString($ItemNumber)
  $lookup = Invoke-ArenaJson -Uri "$apiRoot/items?number=$pnEnc&limit=1" -Headers $headers
  if (-not $lookup.results -or $lookup.results.Count -eq 0) {
    return @{ Found = $false; ItemNumber = $ItemNumber }
  }
  $anyGuid = $lookup.results[0].guid
  $revs = Invoke-ArenaJson -Uri "$apiRoot/items/$anyGuid/revisions" -Headers $headers
  if (-not $revs.results -or $revs.results.Count -eq 0) {
    return @{ Found = $true; ItemNumber = $ItemNumber; Working = $null; CheckedIn = $null }
  }
  $working = $revs.results | Where-Object { $_.status -eq 0 } | Select-Object -First 1
  $checkedIn = @($revs.results | Where-Object { $_.status -eq 1 }) |
    Sort-Object { $_.effectiveDateTime } -Descending |
    Select-Object -First 1
  return @{
    Found     = $true
    ItemNumber = $ItemNumber
    Working   = $working
    CheckedIn = $checkedIn
  }
}

function Get-FilesOnRev {
  param([string]$RevGuid)
  if ([string]::IsNullOrWhiteSpace($RevGuid)) { return @() }
  $resp = Invoke-ArenaJson -Uri "$apiRoot/items/$RevGuid/files" -Headers $headers
  if (-not $resp.results) { return @() }
  return @($resp.results)
}

function Get-BomOnRev {
  param([string]$RevGuid)
  if ([string]::IsNullOrWhiteSpace($RevGuid)) { return @() }
  $resp = Invoke-ArenaJson -Uri "$apiRoot/items/$RevGuid/bom" -Headers $headers
  if (-not $resp.results) { return @() }
  return @($resp.results)
}

function Get-RevItemSnapshot {
  param([string]$RevGuid)
  if ([string]::IsNullOrWhiteSpace($RevGuid)) { return $null }
  $viewEnc = [Uri]::EscapeDataString('full')
  return Invoke-ArenaJson -Uri "$apiRoot/items/${RevGuid}?responseview=$viewEnc" -Headers $headers
}

function Normalize-Text([object]$Value) {
  if ($null -eq $Value) { return '' }
  return ("$Value").Trim()
}

function Split-RefDes([string]$raw) {
  if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
  return @(
    $raw -split '[,\s;]+' |
      ForEach-Object { $_.Trim().ToUpperInvariant() } |
      Where-Object { $_ }
  )
}

function Format-RefDesSet([string]$raw) {
  $set = Split-RefDes -raw $raw
  if ($set.Count -eq 0) { return '(none)' }
  return ($set | Sort-Object) -join ', '
}

function Get-FileRows {
  param([object[]]$Files)
  $rows = New-Object System.Collections.Generic.List[object]
  foreach ($a in @($Files)) {
    if (-not $a -or -not $a.file) { continue }
    $cat = if ($a.file.category -and $a.file.category.name) {
      Normalize-Text $a.file.category.name
    } else { '(no category)' }
    [void]$rows.Add([pscustomobject]@{
        Key     = $cat.ToLowerInvariant()
        Category = $cat
        Title   = Normalize-Text $a.file.title
        Edition = Normalize-Text $a.file.edition
        Number  = Normalize-Text $a.file.number
        Guid    = Normalize-Text $a.file.guid
      })
  }
  return @($rows | Sort-Object { "$($_.Key)|$($_.Title)|$($_.Edition)|$($_.Number)" })
}

function Add-FileRowToBucket {
  param(
    [hashtable]$Table,
    [string]$Key,
    [object]$Row
  )
  if (-not $Table.ContainsKey($Key)) {
    $Table[$Key] = @()
  }
  $Table[$Key] = @($Table[$Key] + $Row)
}

function Get-FileRowsForKey {
  param(
    [hashtable]$Table,
    [string]$Key
  )
  if (-not $Table.ContainsKey($Key)) { return @() }
  return @($Table[$Key])
}

function Compare-FileSets {
  param(
    [object[]]$WorkingFiles,
    [object[]]$CheckedInFiles,
    [string]$Label
  )
  $wRows = Get-FileRows -Files $WorkingFiles
  $cRows = Get-FileRows -Files $CheckedInFiles

  $wByKey = @{}
  foreach ($r in $wRows) {
    Add-FileRowToBucket -Table $wByKey -Key $r.Key -Row $r
  }
  $cByKey = @{}
  foreach ($r in $cRows) {
    Add-FileRowToBucket -Table $cByKey -Key $r.Key -Row $r
  }

  $allKeys = @($wByKey.Keys + $cByKey.Keys | Sort-Object -Unique)
  $found = $false
  foreach ($key in $allKeys) {
    $wList = Get-FileRowsForKey -Table $wByKey -Key $key
    $cList = Get-FileRowsForKey -Table $cByKey -Key $key
    if ($wList.Count -eq 0 -and $cList.Count -gt 0) {
      $found = $true
      foreach ($f in $cList) {
        Write-Output ("    REMOVED on working: [{0}] '{1}' edition={2} ({3})" -f $f.Category, $f.Title, $f.Edition, $f.Number)
        [void]$script:differences.Add("$Label file removed on working: [$($f.Category)] '$($f.Title)' (was on checked-in).")
      }
      continue
    }
    if ($cList.Count -eq 0 -and $wList.Count -gt 0) {
      $found = $true
      foreach ($f in $wList) {
        Write-Output ("    ADDED on working:   [{0}] '{1}' edition={2} ({3})" -f $f.Category, $f.Title, $f.Edition, $f.Number)
        [void]$script:differences.Add("$Label file added on working: [$($f.Category)] '$($f.Title)' (not on checked-in).")
      }
      continue
    }

    $wSig = ($wList | ForEach-Object { "$($_.Title)|$($_.Edition)|$($_.Number)|$($_.Guid)" } | Sort-Object) -join ';;'
    $cSig = ($cList | ForEach-Object { "$($_.Title)|$($_.Edition)|$($_.Number)|$($_.Guid)" } | Sort-Object) -join ';;'
    if ($wSig -ne $cSig) {
      $found = $true
      Write-Output ("    CHANGED [{0}]:" -f ($wList[0].Category))
      Write-Output ("      working:    {0}" -f (($wList | ForEach-Object { "'$($_.Title)' ed=$($_.Edition) $($_.Number)" }) -join '; '))
      Write-Output ("      checked-in: {0}" -f (($cList | ForEach-Object { "'$($_.Title)' ed=$($_.Edition) $($_.Number)" }) -join '; '))
      [void]$script:differences.Add("$Label file changed: [$($wList[0].Category)] working vs checked-in titles/editions differ.")
    }
  }
  if (-not $found) {
    Write-Output '    (files match)'
  }
}

function Get-BomMap {
  param([object[]]$BomLines)
  $map = @{}
  foreach ($line in @($BomLines)) {
    if (-not $line -or -not $line.item -or -not $line.item.number) { continue }
    $pn = Normalize-Text $line.item.number
    if (-not $pn) { continue }
    $key = $pn.ToLowerInvariant()
    if (-not $map.ContainsKey($key)) {
      $map[$key] = [pscustomobject]@{
        PartNumber = $pn
        Quantity   = if ($null -ne $line.quantity) { [string]$line.quantity } else { '' }
        RefDes     = Normalize-Text $line.refDes
        Notes      = Normalize-Text $line.notes
        LineNumber = if ($null -ne $line.lineNumber) { [string]$line.lineNumber } else { '' }
        Name       = Normalize-Text $line.item.name
      }
    }
  }
  return $map
}

function Compare-BomMaps {
  param(
    [hashtable]$WorkingMap,
    [hashtable]$CheckedInMap,
    [string]$Label
  )
  $allPn = @($WorkingMap.Keys + $CheckedInMap.Keys | Sort-Object -Unique)
  $found = $false
  foreach ($key in $allPn) {
    $w = if ($WorkingMap.ContainsKey($key)) { $WorkingMap[$key] } else { $null }
    $c = if ($CheckedInMap.ContainsKey($key)) { $CheckedInMap[$key] } else { $null }
    if (-not $w) {
      $found = $true
      Write-Output ("    REMOVED on working: {0} qty={1} refDes={2}" -f $c.PartNumber, $c.Quantity, (Format-RefDesSet $c.RefDes))
      [void]$script:differences.Add("$Label BOM line removed on working: $($c.PartNumber).")
      continue
    }
    if (-not $c) {
      $found = $true
      Write-Output ("    ADDED on working:   {0} qty={1} refDes={2}" -f $w.PartNumber, $w.Quantity, (Format-RefDesSet $w.RefDes))
      [void]$script:differences.Add("$Label BOM line added on working: $($w.PartNumber).")
      continue
    }
    $lineDiffs = @()
    if ($w.Quantity -ne $c.Quantity) {
      $lineDiffs += "qty working=$($w.Quantity) checked-in=$($c.Quantity)"
    }
    $wSet = New-Object System.Collections.Generic.HashSet[string]
    foreach ($d in (Split-RefDes $w.RefDes)) { [void]$wSet.Add($d) }
    $cSet = New-Object System.Collections.Generic.HashSet[string]
    foreach ($d in (Split-RefDes $c.RefDes)) { [void]$cSet.Add($d) }
    $onlyW = @((Split-RefDes $w.RefDes) | Where-Object { -not $cSet.Contains($_) })
    $onlyC = @((Split-RefDes $c.RefDes) | Where-Object { -not $wSet.Contains($_) })
    if ($onlyW.Count -gt 0) { $lineDiffs += "refDes only on working: [$($onlyW -join ', ')]" }
    if ($onlyC.Count -gt 0) { $lineDiffs += "refDes only on checked-in: [$($onlyC -join ', ')]" }
    if ($w.Notes -ne $c.Notes) {
      $lineDiffs += "notes working='$($w.Notes)' checked-in='$($c.Notes)'"
    }
    if ($lineDiffs.Count -gt 0) {
      $found = $true
      Write-Output ("    CHANGED {0}:" -f $w.PartNumber)
      foreach ($d in $lineDiffs) { Write-Output ("      {0}" -f $d) }
      [void]$script:differences.Add("$Label BOM line changed: $($w.PartNumber) ($($lineDiffs -join '; ')).")
    }
  }
  if (-not $found) {
    Write-Output '    (BOM lines match)'
  }
}

function Compare-RevItemFields {
  param(
    $WorkingRev,
    $CheckedInRev,
    [string]$Label
  )
  $wSnap = Get-RevItemSnapshot -RevGuid $WorkingRev.guid
  $cSnap = Get-RevItemSnapshot -RevGuid $CheckedInRev.guid
  if (-not $wSnap -or -not $cSnap) {
    Write-Output '    SKIP: could not load revision item snapshots.'
    return
  }

  $fields = @(
    @{ Name = 'name';           W = Normalize-Text $wSnap.name;           C = Normalize-Text $cSnap.name }
    @{ Name = 'description';    W = Normalize-Text $wSnap.description;    C = Normalize-Text $cSnap.description }
    @{ Name = 'uom';            W = Normalize-Text $wSnap.uom;            C = Normalize-Text $cSnap.uom }
    @{
      Name = 'lifecyclePhase'
      W    = if ($wSnap.lifecyclePhase) { Normalize-Text $wSnap.lifecyclePhase.name } else { '' }
      C    = if ($cSnap.lifecyclePhase) { Normalize-Text $cSnap.lifecyclePhase.name } else { '' }
    }
  )

  $found = $false
  foreach ($f in $fields) {
    if ($f.W -eq $f.C) { continue }
    $found = $true
    Write-Output ("    CHANGED {0}: working='{1}' checked-in='{2}'" -f $f.Name, $f.W, $f.C)
    [void]$script:differences.Add("$Label $($f.Name) changed: working='$($f.W)' vs checked-in='$($f.C)'.")
  }
  if (-not $found) {
    Write-Output '    (revision item fields match)'
  }
}

function Compare-ItemPair {
  param(
    [string]$Label,
    [string]$ItemNumber
  )
  Write-Output ("=== {0}: {1} ===" -f $Label, $ItemNumber)
  $info = Get-WorkingAndCheckedInRevisions -ItemNumber $ItemNumber
  if (-not $info.Found) {
    Write-Output '  MISSING: item not found in Arena workspace.'
    [void]$differences.Add("$Label item $ItemNumber not found.")
    Write-Output ''
    return
  }
  if (-not $info.Working) {
    Write-Output '  MISSING: no working revision (status 0).'
    [void]$differences.Add("$Label item $ItemNumber has no working revision.")
    Write-Output ''
    return
  }
  if (-not $info.CheckedIn) {
    Write-Output '  MISSING: no checked-in revision (status 1). Nothing to compare against.'
    [void]$differences.Add("$Label item $ItemNumber has no checked-in revision to compare.")
    Write-Output ''
    return
  }

  $diffStartIndex = $differences.Count
  $deferredBlankWorkingRev = $false

  $wNum = Normalize-Text $info.Working.number
  $cNum = Normalize-Text $info.CheckedIn.number
  Write-Output ("  Working rev:     number={0}  guid={1}" -f $(if ($wNum) { $wNum } else { '(none)' }), $info.Working.guid)
  Write-Output ("  Checked-in rev:  number={0}  guid={1}  effective={2}" -f `
    $(if ($cNum) { $cNum } else { '(none)' }), $info.CheckedIn.guid, $info.CheckedIn.effectiveDateTime)

  if ($wNum -ne $cNum) {
    if ([string]::IsNullOrWhiteSpace($wNum)) {
      $deferredBlankWorkingRev = $true
      Write-Output ("  Working revision number is blank (checked-in={0}); ignored unless other changes exist on this item." -f $cNum)
    } else {
      Write-Output ("  REVISION NUMBER CHANGED: working={0} checked-in={1}" -f $wNum, $cNum)
      [void]$differences.Add("$Label revision number differs: working=$wNum checked-in=$cNum.")
    }
  } else {
    Write-Output '  Revision number: unchanged'
  }

  Write-Output ''
  Write-Output '  [Files]'
  $wFiles = Get-FilesOnRev -RevGuid $info.Working.guid
  $cFiles = Get-FilesOnRev -RevGuid $info.CheckedIn.guid
  Write-Output ("    working: {0} file(s); checked-in: {1} file(s)" -f $wFiles.Count, $cFiles.Count)
  Compare-FileSets -WorkingFiles $wFiles -CheckedInFiles $cFiles -Label $Label

  Write-Output ''
  Write-Output '  [BOM]'
  $wBom = Get-BomOnRev -RevGuid $info.Working.guid
  $cBom = Get-BomOnRev -RevGuid $info.CheckedIn.guid
  Write-Output ("    working: {0} line(s); checked-in: {1} line(s)" -f $wBom.Count, $cBom.Count)
  Compare-BomMaps -WorkingMap (Get-BomMap -BomLines $wBom) -CheckedInMap (Get-BomMap -BomLines $cBom) -Label $Label

  Write-Output ''
  Write-Output '  [Other revision fields]'
  Compare-RevItemFields -WorkingRev $info.Working -CheckedInRev $info.CheckedIn -Label $Label

  if ($deferredBlankWorkingRev) {
    $otherChangesOnItem = ($differences.Count -gt $diffStartIndex)
    if ($otherChangesOnItem) {
      Write-Output ("  REVISION NUMBER: working blank vs checked-in {0}" -f $cNum)
      [void]$differences.Add("$Label revision number differs: working=(blank) checked-in=$cNum.")
    } else {
      Write-Output '  Revision number: blank on working (ignored; no other changes on this item)'
    }
  }

  Write-Output ''
}

Write-Output ("=== Compare working vs checked-in ===" )
Write-Output ("  PCBA part:      {0}" -f $pcbaNumber)
Write-Output ("  Bare PCB part:  {0}  (auto-mapped from PCBA)" -f $bareNumber)
Write-Output ''

Compare-ItemPair -Label 'PCBA' -ItemNumber $pcbaNumber
Compare-ItemPair -Label 'Bare PCB' -ItemNumber $bareNumber

function Write-CompareSummaryLines {
  if ($differences.Count -eq 0) {
    Write-Output ("Working revision matches checked-in for PCBA {0} and bare PCB {1} (no changes detected)." -f $pcbaNumber, $bareNumber)
    return
  }

  $pcbaDiffs = @($differences | Where-Object { $_ -like 'PCBA *' })
  $bareDiffs = @($differences | Where-Object { $_ -like 'Bare PCB *' })
  $otherDiffs = @($differences | Where-Object { $_ -notlike 'PCBA *' -and $_ -notlike 'Bare PCB *' })

  Write-Output ("{0} difference(s) between working and checked-in:" -f $differences.Count)
  if ($pcbaDiffs.Count -gt 0) {
    Write-Output ("  PCBA ({0}):" -f $pcbaNumber)
    foreach ($d in $pcbaDiffs) { Write-Output "    - $($d -replace '^PCBA ', '')" }
  }
  if ($bareDiffs.Count -gt 0) {
    Write-Output ("  Bare PCB ({0}):" -f $bareNumber)
    foreach ($d in $bareDiffs) { Write-Output "    - $($d -replace '^Bare PCB ', '')" }
  }
  foreach ($d in $otherDiffs) { Write-Output "  - $d" }
}

if ($DeferSummary) {
  Write-Output '---BEGIN-COMPARE-SUMMARY---'
  Write-CompareSummaryLines
  Write-Output '---END-COMPARE-SUMMARY---'
} else {
  Write-Output '--- Summary ---'
  Write-CompareSummaryLines
}

if ($differences.Count -eq 0) {
  exit 0
}
exit 2
