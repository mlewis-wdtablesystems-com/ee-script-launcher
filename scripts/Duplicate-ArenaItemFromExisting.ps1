<#
.SYNOPSIS
  Duplicate an Arena item to a new part number (untested): preview name/description or perform create + optional BOM/files.

.DESCRIPTION
  Preview mode (-Mode Preview): logs in, loads the source item by old part number, and prints JSON only:
    { "prefill": { "item_name": "...", "item_description": "..." } }
  for the launcher "Load name & description" button.

  Execute mode (default): creates a new item (POST /items) with edited name/description,
  owner set to the Windows session display name (or -OwnerFullName), and revision number from -NewItemRevision.
  Category, UOM, and additional attributes are derived from the source item. Many workspaces do not allow setting
  top-level number, procurementType, or isAssembly on create (Arena 3031); in that case this script sends numberFormat
  with field values derived from -NewPartNumber. The format is resolved from GET /items (if present), else
  GET /settings/items/categories/{categoryGuid} (walking parent categories when the format is inherited), then
  GET /settings/items/numberformats/{guid}. The items list response may include numberFormat when the full item GET does
  not. Delimiter fields from the
  format definition are omitted on create; other fields map to hyphen-separated segments of the source and new numbers
  (optional zero-padding for numeric segments). Single-field formats use the full new part number.
  Lifecycle phase is not copied — Arena assigns the category default. NUMBER additional attributes are sent as JSON numbers;
  MULTI_SELECT/LINKED_ITEM-style attributes are skipped with a warning.
  Optionally copies BOM lines from GET /items/{guid}/bom and/or re-associates files from the chosen source revision
  to the new item's working revision (POST /items/{revGuid}/files).

  Source revision: -SourceRevisionMode working | effective_latest, or -SourceRevisionOverride with a rev letter
  (e.g. B) to copy from that specific revision.

  Credentials: ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID.

.PARAMETER Mode
  Preview  = JSON prefill only (stdout must be only JSON for launcher discovery).
  Execute  = perform duplicate (default).

.EXAMPLE
  .\Duplicate-ArenaItemFromExisting.ps1 -Mode Preview -OldPartNumber 294-10-001
  .\Duplicate-ArenaItemFromExisting.ps1 -OldPartNumber 294-10-001 -NewPartNumber 294-10-099 -ItemName "..." -ItemDescription "..."
#>
param(
  [Parameter(Mandatory = $true)] [string]$OldPartNumber,
  [string]$NewPartNumber = '',
  [string]$ItemName = '',
  [string]$ItemDescription = '',
  [string]$CopyBom = 'no',
  [string]$CopyFiles = 'no',
  [ValidateSet('working', 'effective_latest')] [string]$SourceRevisionMode = 'working',
  [string]$SourceRevisionOverride = '',
  [string]$NewItemRevision = 'A',
  [string]$OwnerFullName = '',
  [ValidateSet('Preview', 'Execute')] [string]$Mode = 'Execute',
  [string]$Email = $env:ARENA_EMAIL,
  [string]$WorkspaceId = $env:ARENA_WORKSPACE_ID,
  [string]$BaseUrl = 'https://api.arenasolutions.com/v1'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
. (Join-Path $scriptsDir 'ArenaSessionAuthor.ps1')

if ([string]::IsNullOrWhiteSpace($Email)) { throw 'Set ARENA_EMAIL.' }
if ([string]::IsNullOrWhiteSpace($env:ARENA_PASSWORD)) { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($WorkspaceId)) { throw 'Set ARENA_WORKSPACE_ID.' }

if ($Mode -eq 'Preview') {
  $OldPartNumber = $OldPartNumber.Trim()
  if (-not $OldPartNumber) { throw 'OldPartNumber is required.' }
}
else {
  if (-not $NewPartNumber.Trim()) { throw 'NewPartNumber is required for Execute mode.' }
  if (-not $ItemName.Trim()) { throw 'ItemName is required for Execute mode (use Preview to pre-fill).' }
}

if ([string]::IsNullOrWhiteSpace($OwnerFullName)) {
  $OwnerFullName = Get-ArenaWindowsSessionAuthorName
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

function Invoke-ArenaJsonPostWithErrorBody {
  <#
  Same as Invoke-ArenaJson for POST, but on failure appends the response body (Arena errors JSON) to the exception message.
  #>
  param(
    [Parameter(Mandatory = $true)] [string]$Uri,
    [hashtable]$Headers,
    [Parameter(Mandatory = $true)] [string]$Body
  )
  try {
    return Invoke-ArenaJson -Uri $Uri -Method Post -Headers $Headers -Body $Body
  } catch {
    $msg = $_.Exception.Message
    if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
      $edm = $_.ErrorDetails.Message.Trim()
      if ($edm) { $msg = "$msg`n`nArena error response:`n$edm" }
    }
    $ex = $_.Exception
    if ($ex.Response) {
      try {
        $stream = $ex.Response.GetResponseStream()
        if ($null -ne $stream) {
          $reader = New-Object System.IO.StreamReader($stream)
          $detail = $reader.ReadToEnd()
          if ($detail -and $msg -notlike "*$detail*") {
            $msg = "$msg`n`nArena error response:`n$detail"
          }
        }
      } catch {}
    }
    throw $msg
  }
}

function Normalize-ArenaItemUom([string]$raw) {
  $t = ($raw -as [string]).Trim()
  if (-not $t) { return 'Each' }
  if ($t -match '^(?i)ea|each$') { return 'Each' }
  return $t
}

function Build-AdditionalAttributesForItemCreate([object[]]$srcAttrs) {
  <#
  Rebuild additionalAttributes for POST /items: correct JSON types for NUMBER, skip unsupported complex types.
  #>
  $out = New-Object System.Collections.Generic.List[hashtable]
  foreach ($a in @($srcAttrs)) {
    if (-not $a -or -not $a.guid) { continue }
    if ($null -eq $a.value) { continue }
    $g = "$($a.guid)".Trim()
    if (-not $g) { continue }
    $ft = ("$($a.fieldType)").Trim().ToUpperInvariant()
    switch -Regex ($ft) {
      '^(NUMBER|INTEGER|FLOAT|CURRENCY)$' {
        $dv = 0.0
        if ([double]::TryParse([string]$a.value, [ref]$dv)) {
          $out.Add(@{ guid = $g; value = $dv })
        }
        elseif ($a.value -is [double] -or $a.value -is [decimal] -or $a.value -is [int]) {
          $out.Add(@{ guid = $g; value = [double]$a.value })
        }
        else {
          Write-Warning "Skipping NUMBER-like attribute (unparseable value): $($a.name)"
        }
      }
      '^CHECKBOX$' {
        $b = $false
        if ($a.value -is [bool]) {
          $out.Add(@{ guid = $g; value = [bool]$a.value })
        }
        elseif ([bool]::TryParse([string]$a.value, [ref]$b)) {
          $out.Add(@{ guid = $g; value = $b })
        }
      }
      '^(MULTI_SELECT|LINKED_ITEM|ITEM_TO_ITEM|CATEGORY)$' {
        Write-Warning "Skipping additional attribute on create (type not supported here): $($a.name) ($ft)"
      }
      default {
        $sv = "$($a.value)".Trim()
        if ($sv.Length -eq 0) { continue }
        $out.Add(@{ guid = $g; value = $sv })
      }
    }
  }
  return @($out)
}

function Format-NumberFieldValueLikeSource([string]$OldFieldValue, [string]$NewSegment) {
  $o = ($OldFieldValue -as [string]).Trim()
  $n = ($NewSegment -as [string]).Trim()
  if ($o.Length -gt $n.Length -and $o -match '^\d+$' -and $n -match '^\d+$') {
    return $n.PadLeft($o.Length, '0')
  }
  return $n
}

function Get-ArenaItemNumberFormatDefinition {
  <#
  Returns number format definition (guid + ordered fields, usually with type) for POST /items.
  GET /items often omits numberFormat. Category settings may inherit the format from a parent category (no numberFormat on
  the leaf). GET /settings/items/numberformats/<guid> is the canonical field list; some workspaces still embed usable
  field guids on GET /settings/items/categories/<guid>.
  #>
  param(
    [Parameter(Mandatory)][object]$SourceItem,
    [Parameter(Mandatory)][hashtable]$Headers,
    [Parameter(Mandatory)][string]$ApiRoot
  )
  $nfGuid = $null
  $embeddedFields = $null

  if ($SourceItem.numberFormat -and "$($SourceItem.numberFormat.guid)".Trim()) {
    $nfGuid = "$($SourceItem.numberFormat.guid)".Trim()
    if ($SourceItem.numberFormat.fields -and @($SourceItem.numberFormat.fields).Count -gt 0) {
      $embeddedFields = @($SourceItem.numberFormat.fields)
    }
  }

  if (-not $nfGuid) {
    $cg = "$($SourceItem.category.guid)".Trim()
    if (-not $cg) { return $null }
    $guard = 0
    while ($cg -and $guard -lt 50) {
      $guard++
      $cat = $null
      try {
        $cat = Invoke-ArenaJson -Uri "$ApiRoot/settings/items/categories/$cg" -Headers $Headers
      } catch {
        $cat = $null
      }
      if (-not $cat) { break }

      if ($cat.numberFormat -and "$($cat.numberFormat.guid)".Trim()) {
        $nfGuid = "$($cat.numberFormat.guid)".Trim()
        if ($cat.numberFormat.fields -and @($cat.numberFormat.fields).Count -gt 0) {
          $embeddedFields = @($cat.numberFormat.fields)
        }
        break
      }

      $pg = $null
      if ($cat.parentCategory) { $pg = "$($cat.parentCategory.guid)".Trim() }
      if (-not $pg) { break }
      $cg = $pg
    }
  }

  if (-not $nfGuid) { return $null }

  $full = $null
  try {
    $full = Invoke-ArenaJson -Uri "$ApiRoot/settings/items/numberformats/$nfGuid" -Headers $Headers
  } catch {
    $full = $null
  }

  if ($full -and $full.errors) {
    $errArr = @($full.errors)
    if ($errArr.Count -gt 0) { $full = $null }
  }

  $fieldList = $null
  $resolvedGuid = "$nfGuid".Trim()
  if ($full) {
    if ("$($full.guid)".Trim()) { $resolvedGuid = "$($full.guid)".Trim() }
    if ($full.fields -and @($full.fields).Count -gt 0) {
      $fieldList = @($full.fields)
    }
    elseif ($full.results -and $full.results[0]) {
      $r0 = $full.results[0]
      if ("$($r0.guid)".Trim()) { $resolvedGuid = "$($r0.guid)".Trim() }
      if ($r0.fields -and @($r0.fields).Count -gt 0) { $fieldList = @($r0.fields) }
    }
  }

  if ($fieldList -and $fieldList.Count -gt 0) {
    return [ordered]@{ guid = $resolvedGuid; fields = $fieldList }
  }

  if ($embeddedFields -and $embeddedFields.Count -gt 0) {
    return [ordered]@{ guid = $resolvedGuid; fields = $embeddedFields }
  }

  return $null
}

function Build-NumberFormatForCreateFromSourceItem {
  <#
  Workspaces that mark core attribute "number" as not creatable (3031) require numberFormat { guid, fields[] }
  on POST /items. Field guids and types come from settings; values come from the source number and -NewPartNumber.
  #>
  param(
    [Parameter(Mandatory)][object]$SourceItem,
    [Parameter(Mandatory)][string]$NewPartNumber,
    [Parameter(Mandatory)][hashtable]$Headers,
    [Parameter(Mandatory)][string]$ApiRoot
  )
  $newPn = $NewPartNumber.Trim()
  $oldPn = ("$($SourceItem.number)").Trim()
  if (-not $oldPn) { throw 'Source item has no number; cannot derive numberFormat field values.' }

  $newSegs = @($newPn -split '-')
  $oldSegs = @($oldPn -split '-')
  if ($newSegs.Count -ne $oldSegs.Count) {
    throw (
      "New part number '$newPn' must have the same hyphen-separated segment count as the source number '$oldPn' " +
      "($($newSegs.Count) vs $($oldSegs.Count))."
    )
  }

  $defs = Get-ArenaItemNumberFormatDefinition -SourceItem $SourceItem -Headers $Headers -ApiRoot $ApiRoot
  if (-not $defs -or -not $defs.guid) { return $null }
  $nfGuid = "$($defs.guid)".Trim()
  $fieldDefs = @($defs.fields)
  if ($fieldDefs.Count -eq 0) { return $null }

  $hasTypes = $false
  foreach ($fd in $fieldDefs) {
    if ("$($fd.type)".Trim()) { $hasTypes = $true; break }
  }

  $outFields = New-Object System.Collections.Generic.List[hashtable]

  if ($hasTypes) {
    $segIdx = 0
    foreach ($fd in $fieldDefs) {
      $t = ("$($fd.type)").Trim().ToUpperInvariant()
      if ($t -eq 'DELIMITER') { continue }
      if ($segIdx -ge $oldSegs.Count) {
        throw (
          "Number format '$nfGuid' has more non-delimiter fields than hyphen segments in '$oldPn' " +
          "(extra field after segment index $segIdx)."
        )
      }
      $gf = "$($fd.guid)".Trim()
      if (-not $gf) { return $null }
      $oldFv = $oldSegs[$segIdx]
      $seg = Format-NumberFieldValueLikeSource $oldFv $newSegs[$segIdx]
      $outFields.Add(@{ guid = $gf; value = $seg })
      $segIdx++
    }
    if ($segIdx -ne $oldSegs.Count) {
      throw (
        "Cannot map '$oldPn' to number format: after skipping DELIMITER fields, " +
        "expected $($oldSegs.Count) value field(s) but walked $segIdx (format guid $nfGuid). " +
        "If the part number does not use hyphens the same way as this format, adjust -NewPartNumber or the source item."
      )
    }
    return [ordered]@{ guid = $nfGuid; fields = @($outFields) }
  }

  if ($fieldDefs.Count -eq 1) {
    $gf = "$($fieldDefs[0].guid)".Trim()
    if (-not $gf) { return $null }
    return [ordered]@{ guid = $nfGuid; fields = @(@{ guid = $gf; value = $newPn }) }
  }

  if ($fieldDefs.Count -eq $oldSegs.Count) {
    for ($i = 0; $i -lt $fieldDefs.Count; $i++) {
      $gf = "$($fieldDefs[$i].guid)".Trim()
      if (-not $gf) { return $null }
      $oldFv = "$($fieldDefs[$i].value)".Trim()
      if (-not $oldFv) { $oldFv = $oldSegs[$i] }
      $seg = Format-NumberFieldValueLikeSource $oldFv $newSegs[$i]
      $outFields.Add(@{ guid = $gf; value = $seg })
    }
    return [ordered]@{ guid = $nfGuid; fields = @($outFields) }
  }

  throw (
    "Cannot map -NewPartNumber to number format: definition has $($fieldDefs.Count) field(s) without type metadata, " +
    "but '$newPn' splits into $($oldSegs.Count) hyphen segment(s) (and is not a single-field format). " +
    "Format guid: $nfGuid."
  )
}

$apiRoot = $BaseUrl.TrimEnd('/')

$loginBody = @{
  email       = $Email
  password    = $env:ARENA_PASSWORD
  workspaceId = [long]$WorkspaceId
} | ConvertTo-Json

$login = Invoke-ArenaJson -Uri "$apiRoot/login" -Method Post -Body $loginBody
$sessionId = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sessionId)) { $sessionId = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sessionId)) { throw 'Login response did not include arenaSessionId.' }

$headers = @{
  arena_session_id = $sessionId
  'Content-Type'   = 'application/json'
}

$pnEnc = [Uri]::EscapeDataString($OldPartNumber.Trim())
$lookup = Invoke-ArenaJson -Uri "$apiRoot/items?number=$pnEnc&limit=1&responseview=full" -Headers $headers
if (-not $lookup.results -or $lookup.results.Count -eq 0) {
  throw "No Arena item found for old part number '$($OldPartNumber.Trim())'."
}
$srcLite = $lookup.results[0]
$srcGuid = "$($srcLite.guid)".Trim()
if (-not $srcGuid) { throw 'Source item has no guid.' }

$src = Invoke-ArenaJson -Uri "$apiRoot/items/$($srcGuid)?responseview=full&includeEmptyAdditionalAttributes=true" -Headers $headers
if (-not $src) { throw "GET /items/$srcGuid returned empty." }

if ($srcLite.numberFormat -and "$($srcLite.numberFormat.guid)".Trim()) {
  if (-not $src.numberFormat -or -not "$($src.numberFormat.guid)".Trim()) {
    $src | Add-Member -MemberType NoteProperty -Name 'numberFormat' -Value $srcLite.numberFormat -Force
  }
}

$srcName = if ($null -ne $src.name) { "$($src.name)".Trim() } else { '' }
$srcDesc = if ($null -ne $src.description) { "$($src.description)".Trim() } else { '' }

if ($Mode -eq 'Preview') {
  $payload = [ordered]@{
    prefill = [ordered]@{
      item_name        = $srcName
      item_description = $srcDesc
    }
    old_part_number = $OldPartNumber.Trim()
  }
  [Console]::Out.WriteLine(($payload | ConvertTo-Json -Depth 6 -Compress))
  exit 0
}

# ---------- Execute ----------
$newPn = $NewPartNumber.Trim()
$pnNewEnc = [Uri]::EscapeDataString($newPn)
$exists = Invoke-ArenaJson -Uri "$apiRoot/items?number=$pnNewEnc&limit=1" -Headers $headers
if ($exists.results -and $exists.results.Count -gt 0) {
  throw "An item with number '$newPn' already exists in Arena. Choose a different new part number."
}

function Select-SourceRevision {
  param([object[]]$RevResults, [string]$ModeKey, [string]$Override)
  if (($Override -ne $null) -and $Override.Trim()) {
    $key = $Override.Trim()
    $m = $RevResults |
      Where-Object { $null -ne $_.number -and ("$($_.number)".Trim() -ieq $key) } |
      Select-Object -First 1
    if (-not $m) {
      $avail = ($RevResults | ForEach-Object { "$($_.number) (status=$($_.status))" }) -join ', '
      throw "No revision matches override '$key'. Available: $avail"
    }
    return $m
  }
  if ($ModeKey -eq 'working') {
    $w = $RevResults | Where-Object { $_.status -eq 0 } | Select-Object -First 1
    if (-not $w) { throw 'No working revision (status 0) on source item.' }
    return $w
  }
  $eff = @($RevResults | Where-Object { $_.status -eq 1 })
  if ($eff.Count -eq 0) { throw 'No effective (released) revision (status 1) on source item for effective_latest mode.' }
  return $eff | Sort-Object { $_.effectiveDateTime } -Descending | Select-Object -First 1
}

$revs = Invoke-ArenaJson -Uri "$apiRoot/items/$srcGuid/revisions" -Headers $headers
if (-not $revs.results -or $revs.results.Count -eq 0) {
  throw "No revisions returned for source item $($OldPartNumber.Trim())."
}
$srcRev = Select-SourceRevision -RevResults @($revs.results) -ModeKey $SourceRevisionMode -Override $SourceRevisionOverride
$srcRevGuid = "$($srcRev.guid)".Trim()
Write-Host ("=== Source revision: number={0} status={1} guid={2} ===" -f $srcRev.number, $srcRev.status, $srcRevGuid)

$doBom = $CopyBom -match '^(?i)yes|1|true$'
$doFiles = $CopyFiles -match '^(?i)yes|1|true$'

$numberFormatBody = Build-NumberFormatForCreateFromSourceItem -SourceItem $src -NewPartNumber $newPn -Headers $headers -ApiRoot $apiRoot
if (-not $numberFormatBody) {
  throw (
    "This workspace does not allow setting item 'number' directly on create (Arena 3031). " +
    "Could not resolve a number format: no usable numberFormat on GET /items or the items lookup, " +
    "and no field list from GET /settings/items/categories (including parent categories) or GET /settings/items/numberformats. " +
    "Category guid: $($src.category.guid). Confirm this category (or an ancestor) has an item number format in Arena settings."
  )
}

$create = [ordered]@{
  category       = @{ guid = "$($src.category.guid)" }
  name           = $ItemName.Trim()
  description    = $ItemDescription.Trim()
  revisionNumber = $(if ($NewItemRevision.Trim()) { $NewItemRevision.Trim() } else { 'A' })
  owner          = @{ fullName = $OwnerFullName.Trim() }
  uom            = (Normalize-ArenaItemUom $(if ($src.uom) { "$($src.uom)" } else { 'Each' }))
  numberFormat   = $numberFormatBody
}
if ($src.additionalAttributes -and @($src.additionalAttributes).Count -gt 0) {
  $attrs = @(Build-AdditionalAttributesForItemCreate @($src.additionalAttributes))
  if ($attrs.Count -gt 0) { $create.additionalAttributes = $attrs }
}

$createJson = $create | ConvertTo-Json -Depth 12
Write-Host '=== POST /items (create duplicate) ==='
Write-Host ('  Request body (truncated): {0}' -f ($createJson.Substring(0, [Math]::Min(2000, $createJson.Length))))
$created = Invoke-ArenaJsonPostWithErrorBody -Uri "$apiRoot/items" -Headers $headers -Body $createJson
$newGuid = if ($created.guid) { "$($created.guid)".Trim() } else { '' }
if (-not $newGuid) {
  if ($created.errors) {
    throw ("Arena rejected create: " + (($created.errors | ConvertTo-Json -Compress)))
  }
  throw "Create response did not include item guid. Check workspace rules (number format, required attributes)."
}
Write-Host ("  New item guid: {0}" -f $newGuid)
Write-Host ("  App URL:       {0}" -f $(if ($created.url.app) { $created.url.app } else { '(see Arena)' }))

$newRevs = Invoke-ArenaJson -Uri "$apiRoot/items/$newGuid/revisions" -Headers $headers
$newWorking = $newRevs.results | Where-Object { $_.status -eq 0 } | Select-Object -First 1
if (-not $newWorking) { throw 'New item has no working revision (status 0).' }
$newWorkingRevGuid = "$($newWorking.guid)".Trim()

if ($doBom) {
  Write-Host '=== Copy BOM lines ==='
  try {
    $bom = Invoke-ArenaJson -Uri "$apiRoot/items/$srcGuid/bom" -Headers $headers
  } catch {
    Write-Warning "Could not read BOM on source: $($_.Exception.Message)"
    $bom = $null
  }
  if ($bom -and $bom.results) {
    $n = 0
    foreach ($line in @($bom.results)) {
      if (-not $line.item -or -not $line.item.guid) { continue }
      $lineBody = [ordered]@{
        refDes   = $line.refDes
        quantity = $line.quantity
        notes    = $line.notes
        item     = @{ guid = "$($line.item.guid)" }
      }
      if ($line.additionalAttributes -and @($line.additionalAttributes).Count -gt 0) {
        $la = @()
        foreach ($x in @($line.additionalAttributes)) {
          if (-not $x.guid) { continue }
          $la += @{ guid = "$($x.guid)"; value = $x.value }
        }
        if ($la.Count -gt 0) { $lineBody.additionalAttributes = $la }
      }
      $lj = $lineBody | ConvertTo-Json -Depth 8
      try {
        [void](Invoke-ArenaJson -Uri "$apiRoot/items/$newGuid/bom" -Method Post -Headers $headers -Body $lj)
        $n++
      } catch {
        Write-Warning "BOM line add failed (refDes=$($line.refDes)): $($_.Exception.Message)"
      }
    }
    Write-Host ("  Added {0} BOM line(s)." -f $n)
  } else {
    Write-Host '  (No BOM lines on source or not an assembly.)'
  }
}

if ($doFiles) {
  Write-Host '=== Copy file associations ==='
  try {
    $filesResp = Invoke-ArenaJson -Uri "$apiRoot/items/$srcRevGuid/files" -Headers $headers
  } catch {
    Write-Warning "Could not list files on source revision: $($_.Exception.Message)"
    $filesResp = $null
  }
  $m = 0
  if ($filesResp -and $filesResp.results) {
    foreach ($a in @($filesResp.results)) {
      if (-not $a.file -or -not $a.file.guid) { continue }
      if ($a.file.locked -eq $true) {
        Write-Warning "Skipping locked file $($a.file.title)."
        continue
      }
      $fg = "$($a.file.guid)".Trim()
      $assocBody = @{
        latestEditionAssociation = $true
        primary                  = $false
        file                     = @{ guid = $fg }
      } | ConvertTo-Json
      try {
        [void](Invoke-ArenaJson -Uri "$apiRoot/items/$newWorkingRevGuid/files" -Method Post -Headers $headers -Body $assocBody)
        $m++
      } catch {
        Write-Warning "File association failed ($($a.file.title)): $($_.Exception.Message)"
      }
    }
  }
  Write-Host ("  Associated {0} file(s) on new working revision." -f $m)
}

Write-Host ''
Write-Host 'Done. Duplicate item flow finished (untested — verify in Arena).' -ForegroundColor Green
exit 0
