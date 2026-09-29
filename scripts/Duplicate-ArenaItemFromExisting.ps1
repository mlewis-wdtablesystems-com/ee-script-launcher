<#
.SYNOPSIS
  Duplicate an Arena item to a new part number: preview name/description or perform create + optional BOM/files.

.DESCRIPTION
  Preview mode (-Mode Preview): logs in, loads the source item by old part number, and prints JSON only:
    { "prefill": { "item_name": "...", "item_description": "..." } }
  for the launcher "Load name & description" button.

  Execute mode (default): creates a new item (POST /items) with edited name/description,
  owner set to the Windows session display name (or -OwnerFullName), and revision number from -NewItemRevision.
  If Item description is left blank, the source item description is used; if that is also empty, the item name is
  used (Arena rejects an empty description on create). Category, UOM, and additional attributes are derived from the source item. Many workspaces do not allow setting
  top-level number, procurementType, or isAssembly on create (Arena 3031); in that case this script sends numberFormat
  with field values derived from -NewPartNumber. The format is resolved from GET /items (if present), else
  GET /settings/items/categories/{categoryGuid} (walking parent categories when the format is inherited), then
  GET /settings/items/numberformats/{guid}. When those omit a format (common for PCBA categories), the script infers
  the format by matching both the source and new part numbers against workspace number formats (VALUE_LIST codes,
  segment shape, or a single FREE_TEXT field such as Basic Item Number). When several formats match the source number,
  the new number must also validate (e.g. WDTS P/N project codes); prefer the most specific matching format.
  AUTO_SEQUENCE fields (WDTS P/N last segment) are omitted on create — Arena assigns the next sequence and ignores
  any value sent for that field. After create, if the assigned number is not exactly -NewPartNumber, the script
  renumbers via the workspace FREE_TEXT "Basic Item Number" format and fails if that still does not match.
  The items list response may include numberFormat when the full item GET does
  not. Delimiter fields from the
  format definition are omitted on create; other fields map to hyphen-separated segments of the source and new numbers
  (optional zero-padding for numeric segments). Single-field formats use the full new part number.
  Lifecycle phase is not copied — Arena assigns the category default. NUMBER additional attributes are sent as JSON numbers;
  MULTI_SELECT/LINKED_ITEM-style attributes are skipped with a warning.
  Optionally copies BOM lines from GET /items/{guid}/bom and/or re-associates files from the chosen source revision
  to the new item's working revision (POST /items/{revGuid}/files).

  Source revision: -SourceRevisionMode working | effective_latest, or -SourceRevisionOverride with a rev letter
  (e.g. B) to copy from that specific revision.

  After create, the new item gets the Arena display image (Illustrative Image / Product Image):
    -UseProductImage yes (default) copies the source revision's Product Image file association;
    otherwise -ItemImageFilePath uploads that local file (title "Product Image <newPn> Rev <letter>").
  -ItemImageFileGuid associates an already-uploaded file (used when duplicating the mapped bare board
  so it receives the same linked image).

  When the source is a PCBA (PPP-10-XXX or PPP-15-XXX), optionally after a successful duplicate the same flow runs
  for the mapped bare board (PPP-20 / PPP-25, or legacy PPP-10-XXX-PCB / PPP-15-XXX-PCB) when -AddBarePcbAlso is yes. Bare board name and description match the
  new PCBA; any "PCBA" in that text is changed to "PCB". -NewItemRevision uses the letter only (e.g. A1 -> A).
  The bare board gets the same display-image file as the new PCBA.

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
  [string]$AddBarePcbAlso = 'no',
  [string]$UseProductImage = 'yes',
  [string]$ItemImageFilePath = '',
  [string]$ItemImageFileGuid = '',
  [switch]$BarePcbFollowUp,
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
  Same as Invoke-ArenaJson for POST/PUT, but on failure appends the response body (Arena errors JSON) to the exception message.
  #>
  param(
    [Parameter(Mandatory = $true)] [string]$Uri,
    [hashtable]$Headers,
    [Parameter(Mandatory = $true)] [string]$Body,
    [ValidateSet('Post', 'Put')] [string]$Method = 'Post'
  )
  try {
    return Invoke-ArenaJson -Uri $Uri -Method $Method -Headers $Headers -Body $Body
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

function Get-DuplicateRevLetterOnly([string]$rev) {
  if ([string]::IsNullOrWhiteSpace($rev)) { return 'A' }
  $m = [regex]::Match($rev.Trim(), '^[A-Za-z]+')
  if ($m.Success) { return $m.Value.ToUpperInvariant() }
  return $rev.Trim()
}

function Test-IsIllustrativeImageCategory {
  param([object]$Association)
  if (-not $Association -or -not $Association.file -or -not $Association.file.category) { return $false }
  $cat = $Association.file.category
  $n = $cat.name
  if ($null -ne $n -and $n.ToString().Trim().ToLowerInvariant() -eq 'illustrative image') {
    return $true
  }
  $path = $cat.path
  if ($null -ne $path -and $path.ToString().ToLowerInvariant().EndsWith('\illustrative image')) {
    return $true
  }
  return $false
}

function Test-IsProductImageTitle([string]$Title) {
  if ([string]::IsNullOrWhiteSpace($Title)) { return $false }
  return ($Title.Trim() -match '(?i)^Product Image\b')
}

function Select-ProductImageAssociation {
  param([object]$FilesResp)
  if (-not $FilesResp -or -not $FilesResp.results) { return $null }
  $rows = @($FilesResp.results | Where-Object { $_.file -and $_.file.guid })
  $illus = @($rows | Where-Object { Test-IsIllustrativeImageCategory $_ })
  $named = @($illus | Where-Object { Test-IsProductImageTitle ("$($_.file.title)") })
  if ($named.Count -gt 0) { return $named[0] }
  if ($illus.Count -gt 0) { return $illus[0] }
  return $null
}

function Test-RevisionHasFileGuid {
  param(
    [Parameter(Mandatory = $true)] [string]$RevGuid,
    [Parameter(Mandatory = $true)] [string]$FileGuid,
    [hashtable]$Headers,
    [string]$ApiRoot
  )
  $filesResp = Invoke-ArenaJson -Uri "$ApiRoot/items/$RevGuid/files" -Headers $Headers
  if (-not $filesResp -or -not $filesResp.results) { return $false }
  $want = $FileGuid.Trim()
  foreach ($a in @($filesResp.results)) {
    if (-not $a.file -or -not $a.file.guid) { continue }
    if (("$($a.file.guid)").Trim() -eq $want) { return $true }
  }
  return $false
}

function Add-ArenaRevisionFileAssociation {
  param(
    [Parameter(Mandatory = $true)] [string]$RevGuid,
    [Parameter(Mandatory = $true)] [string]$FileGuid,
    [hashtable]$Headers,
    [string]$ApiRoot
  )
  $fg = $FileGuid.Trim()
  if (-not $fg) { throw 'File guid is empty; cannot associate item image.' }
  if (Test-RevisionHasFileGuid -RevGuid $RevGuid -FileGuid $fg -Headers $Headers -ApiRoot $ApiRoot) {
    Write-Host ("  Item image already associated (file guid {0})." -f $fg)
    return
  }
  $assocBody = @{
    latestEditionAssociation = $true
    primary                  = $false
    file                     = @{ guid = $fg }
  } | ConvertTo-Json
  [void](Invoke-ArenaJson -Uri "$ApiRoot/items/$RevGuid/files" -Method Post -Headers $Headers -Body $assocBody)
  Write-Host ("  Associated item image file guid {0}." -f $fg)
}

function Resolve-IllustrativeImageCategoryGuid {
  param(
    [hashtable]$Headers,
    [string]$ApiRoot
  )
  $cats = Invoke-ArenaJson -Uri "$ApiRoot/settings/files/categories" -Headers $Headers
  if (-not $cats.results) { throw 'Could not fetch file categories.' }
  $catKey = 'illustrative image'
  $matchedCat = $cats.results | Where-Object {
    $_.assignable -eq $true -and (
      ($_.name -ne $null -and $_.name.ToString().ToLowerInvariant() -eq $catKey) -or
      ($_.path -ne $null -and $_.path.ToString().ToLowerInvariant() -eq $catKey)
    )
  } | Select-Object -First 1
  if (-not $matchedCat) {
    $matchedCat = $cats.results | Where-Object {
      $_.assignable -eq $true -and $_.path -ne $null -and $_.path.ToString().ToLowerInvariant().EndsWith("\$catKey")
    } | Select-Object -First 1
  }
  if (-not $matchedCat) {
    throw "File category 'Illustrative Image' was not found in this workspace."
  }
  return "$($matchedCat.guid)".Trim()
}

function New-ArenaFileFromLocalPath {
  param(
    [Parameter(Mandatory = $true)] [string]$FilePath,
    [Parameter(Mandatory = $true)] [string]$Title,
    [Parameter(Mandatory = $true)] [string]$CategoryGuid,
    [Parameter(Mandatory = $true)] [string]$SessionId,
    [Parameter(Mandatory = $true)] [string]$ApiRoot,
    [string]$Author = ''
  )
  if (-not (Test-Path -LiteralPath $FilePath)) { throw "Item image file not found: $FilePath" }
  Add-Type -AssemblyName System.Net.Http -ErrorAction Stop
  $fileInfo = Get-Item -LiteralPath $FilePath
  $fileName = $fileInfo.Name
  $ext = [System.IO.Path]::GetExtension($fileName).TrimStart('.')
  if ([string]::IsNullOrWhiteSpace($ext)) { $ext = 'bin' }
  $boundary = [System.Guid]::NewGuid().ToString('N')
  $enc = [System.Text.Encoding]::UTF8
  $LF = "`r`n"
  $bodyStream = New-Object System.IO.MemoryStream

  function Write-StringToStream {
    param([System.IO.MemoryStream]$Stream, [string]$Text)
    $b = $enc.GetBytes($Text)
    $Stream.Write($b, 0, $b.Length)
  }
  function Add-MultipartField {
    param([System.IO.MemoryStream]$Stream, [string]$Boundary, [string]$Name, [string]$Value)
    if ($null -eq $Value) { $Value = '' }
    $part = "--$Boundary$LF" +
            "Content-Disposition: form-data; name=`"$Name`"$LF$LF" +
            "$Value$LF"
    Write-StringToStream -Stream $Stream -Text $part
  }
  function Add-MultipartFile {
    param(
      [System.IO.MemoryStream]$Stream,
      [string]$Boundary,
      [string]$Name,
      [string]$FileName,
      [byte[]]$Bytes
    )
    $header = "--$Boundary$LF" +
              "Content-Disposition: form-data; name=`"$Name`"; filename=`"$FileName`"$LF" +
              "Content-Type: application/octet-stream$LF$LF"
    Write-StringToStream -Stream $Stream -Text $header
    $Stream.Write($Bytes, 0, $Bytes.Length)
    Write-StringToStream -Stream $Stream -Text $LF
  }

  Add-MultipartField $bodyStream $boundary 'storageMethodName' 'FILE'
  Add-MultipartField $bodyStream $boundary 'category.guid'     $CategoryGuid
  Add-MultipartField $bodyStream $boundary 'title'             $Title
  Add-MultipartField $bodyStream $boundary 'edition'           '1'
  Add-MultipartField $bodyStream $boundary 'format'            $ext
  Add-MultipartField $bodyStream $boundary 'private'           'false'
  if (-not [string]::IsNullOrWhiteSpace($Author)) {
    Add-MultipartField $bodyStream $boundary 'author.fullName' $Author
  }
  $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
  Add-MultipartFile $bodyStream $boundary 'content' $fileName $fileBytes
  Write-StringToStream -Stream $bodyStream -Text ("--$boundary--$LF")
  $bodyBytes = $bodyStream.ToArray()

  $httpClient = New-Object System.Net.Http.HttpClient
  $httpClient.Timeout = [TimeSpan]::FromMinutes(5)
  $httpClient.DefaultRequestHeaders.Add('arena_session_id', $SessionId)
  $reqContent = New-Object System.Net.Http.ByteArrayContent (,$bodyBytes)
  $reqContent.Headers.ContentType =
    [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("multipart/form-data; boundary=$boundary")
  try {
    $resp = $httpClient.PostAsync("$ApiRoot/files", $reqContent).GetAwaiter().GetResult()
  } catch {
    throw "Arena POST /files transport error: $($_.Exception.Message)"
  }
  $respBody = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
  if (-not $resp.IsSuccessStatusCode) {
    throw ("Arena POST /files returned {0}: {1}" -f [int]$resp.StatusCode, $respBody)
  }
  $fileGuid = $null
  if ($resp.Headers.Contains('X-Arena-FileGuid')) {
    $fileGuid = ($resp.Headers.GetValues('X-Arena-FileGuid') | Select-Object -First 1)
  }
  if ([string]::IsNullOrWhiteSpace($fileGuid) -and -not [string]::IsNullOrWhiteSpace($respBody)) {
    try { $fileGuid = ($respBody | ConvertFrom-Json).guid } catch {}
  }
  if ([string]::IsNullOrWhiteSpace($fileGuid)) {
    throw "File upload returned no file GUID. Body: $respBody"
  }
  Write-Host ("  Uploaded item image '{0}' ({1} bytes) -> {2}" -f $fileName, $fileInfo.Length, $fileGuid)
  return "$fileGuid".Trim()
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

function Get-ArenaNumberFormatValueFields {
  param([object]$FieldDefs)
  $out = @()
  foreach ($fd in @($FieldDefs)) {
    if (-not $fd) { continue }
    $t = ("$($fd.type)").Trim().ToUpperInvariant()
    if ($t -eq 'DELIMITER') { continue }
    $out += $fd
  }
  return @($out)
}

function Get-ArenaBasicItemNumberFormat {
  <#
  Prefer workspace "Basic Item Number" (single FREE_TEXT field) for forcing an exact part number.
  Arena WDTS P/N uses AUTO_SEQUENCE for the last segment and ignores any value sent for that field.
  #>
  param(
    [Parameter(Mandatory)][hashtable]$Headers,
    [Parameter(Mandatory)][string]$ApiRoot
  )
  $list = $null
  try {
    $list = Invoke-ArenaJson -Uri "$ApiRoot/settings/items/numberformats" -Headers $Headers
  } catch {
    return $null
  }
  if (-not $list -or -not $list.results) { return $null }

  $fallback = $null
  foreach ($row in @($list.results)) {
    $g = "$($row.guid)".Trim()
    if (-not $g) { continue }
    $full = $null
    try {
      $full = Invoke-ArenaJson -Uri "$ApiRoot/settings/items/numberformats/$g" -Headers $Headers
    } catch {
      continue
    }
    if (-not $full -or -not $full.fields) { continue }
    $valFields = @(Get-ArenaNumberFormatValueFields -FieldDefs @($full.fields))
    if ($valFields.Count -ne 1) { continue }
    $t = ("$($valFields[0].type)").Trim().ToUpperInvariant()
    if ($t -ne 'FREE_TEXT') { continue }
    $fieldGuid = "$($valFields[0].guid)".Trim()
    if (-not $fieldGuid) { continue }
    $name = if ("$($full.name)".Trim()) { "$($full.name)".Trim() } elseif ("$($row.name)".Trim()) { "$($row.name)".Trim() } else { $g }
    $resolved = [ordered]@{
      guid      = $(if ("$($full.guid)".Trim()) { "$($full.guid)".Trim() } else { $g })
      fieldGuid = $fieldGuid
      name      = $name
    }
    if ($name -ieq 'Basic Item Number') { return $resolved }
    if (-not $fallback) { $fallback = $resolved }
  }
  return $fallback
}

function Set-ArenaItemExactNumber {
  param(
    [Parameter(Mandatory)][string]$ItemGuid,
    [Parameter(Mandatory)][string]$DesiredNumber,
    [Parameter(Mandatory)][hashtable]$Headers,
    [Parameter(Mandatory)][string]$ApiRoot
  )
  $desired = $DesiredNumber.Trim()
  $ft = Get-ArenaBasicItemNumberFormat -Headers $Headers -ApiRoot $ApiRoot
  if (-not $ft) {
    throw (
      "Arena assigned a different item number than requested, and no FREE_TEXT " +
      "'Basic Item Number' format is available to renumber item $ItemGuid to '$desired'."
    )
  }
  $body = @{
    numberFormat = @{
      guid   = $ft.guid
      fields = @(@{ guid = $ft.fieldGuid; value = $desired })
    }
  } | ConvertTo-Json -Depth 6
  Write-Host (
    "  Requested number differs from Arena auto-sequence; renumbering via '$($ft.name)' to '$desired'."
  ) -ForegroundColor DarkYellow
  return (Invoke-ArenaJsonPostWithErrorBody -Uri "$ApiRoot/items/$ItemGuid" -Headers $Headers -Body $body -Method Put)
}

function Get-ArenaNumberFormatPrimaryDelimiter {
  param([object]$FieldDefs)
  foreach ($fd in @($FieldDefs)) {
    $t = ("$($fd.type)").Trim().ToUpperInvariant()
    if ($t -ne 'DELIMITER') { continue }
    $d = ("$($fd.value)").Trim()
    if ($d) { return $d }
  }
  return '-'
}

function Test-PartNumberMatchesArenaNumberFormatFields {
  param(
    [Parameter(Mandatory)][string]$PartNumber,
    [Parameter(Mandatory)][object]$FieldDefs
  )
  $pn = $PartNumber.Trim()
  if (-not $pn) { return $false }
  $valFields = Get-ArenaNumberFormatValueFields -FieldDefs $FieldDefs
  if (@($valFields).Count -eq 0) { return $false }

  if (@($valFields).Count -eq 1) {
    $t = ("$($valFields[0].type)").Trim().ToUpperInvariant()
    if ($t -eq 'FREE_TEXT') { return $true }
  }

  $delim = Get-ArenaNumberFormatPrimaryDelimiter -FieldDefs $FieldDefs
  $segs = @($pn -split [regex]::Escape($delim))
  if ($segs.Count -ne @($valFields).Count) { return $false }

  for ($i = 0; $i -lt @($valFields).Count; $i++) {
    $fd = $valFields[$i]
    $seg = $segs[$i]
    $t = ("$($fd.type)").Trim().ToUpperInvariant()
    if ($t -eq 'VALUE_LIST') {
      $pvs = @($fd.possibleValues)
      if ($pvs.Count -eq 0) { continue }
      $hit = $false
      foreach ($pv in $pvs) {
        if ("$($pv.value)" -eq $seg) { $hit = $true; break }
      }
      if (-not $hit) { return $false }
    }
    elseif ($t -eq 'AUTO_SEQUENCE') {
      # Sequential segment: digits only (optional zero-padding).
      if ($seg -notmatch '^\d+$') { return $false }
    }
  }
  return $true
}

function Get-ArenaNumberFormatMatchSpecificityScore {
  param(
    [Parameter(Mandatory)][string]$PartNumber,
    [Parameter(Mandatory)][object]$FieldDefs
  )
  $valFields = Get-ArenaNumberFormatValueFields -FieldDefs $FieldDefs
  if (@($valFields).Count -eq 0) { return 0 }
  if (@($valFields).Count -eq 1) {
    $t = ("$($valFields[0].type)").Trim().ToUpperInvariant()
    if ($t -eq 'FREE_TEXT') { return 1 }
  }

  $delim = Get-ArenaNumberFormatPrimaryDelimiter -FieldDefs $FieldDefs
  $segs = @($PartNumber.Trim() -split [regex]::Escape($delim))
  if ($segs.Count -ne @($valFields).Count) { return 0 }

  $score = @($valFields).Count
  for ($i = 0; $i -lt @($valFields).Count; $i++) {
    $fd = $valFields[$i]
    $t = ("$($fd.type)").Trim().ToUpperInvariant()
    if ($t -eq 'VALUE_LIST') {
      $pvs = @($fd.possibleValues)
      if ($pvs.Count -eq 0) { continue }
      foreach ($pv in $pvs) {
        if ("$($pv.value)" -eq $segs[$i]) {
          $score += 10
          break
        }
      }
    }
  }
  return $score
}

function Resolve-ArenaNumberFormatDefinitionFromPartNumber {
  <#
  When category/item responses omit numberFormat, infer it by matching part numbers against each workspace format.
  Both the source number and (when provided) the new number must validate — e.g. a new project code not in a
  VALUE_LIST rules out that structured format and leaves Basic Item Number (single FREE_TEXT field).
  #>
  param(
    [Parameter(Mandatory)][string]$PartNumber,
    [string]$NewPartNumber = '',
    [Parameter(Mandatory)][hashtable]$Headers,
    [Parameter(Mandatory)][string]$ApiRoot
  )
  $pn = $PartNumber.Trim()
  if (-not $pn) { return $null }
  $newPn = $NewPartNumber.Trim()

  $list = $null
  try {
    $list = Invoke-ArenaJson -Uri "$ApiRoot/settings/items/numberformats" -Headers $Headers
  } catch {
    return $null
  }
  if (-not $list -or -not $list.results -or @($list.results).Count -eq 0) { return $null }

  $matches = @()
  foreach ($row in @($list.results)) {
    $g = "$($row.guid)".Trim()
    if (-not $g) { continue }
    $full = $null
    try {
      $full = Invoke-ArenaJson -Uri "$ApiRoot/settings/items/numberformats/$g" -Headers $Headers
    } catch {
      continue
    }
    if (-not $full -or ($full.errors -and @($full.errors).Count -gt 0)) { continue }
    $fields = @($full.fields)
    if ($fields.Count -eq 0) { continue }
    if (-not (Test-PartNumberMatchesArenaNumberFormatFields -PartNumber $pn -FieldDefs $fields)) { continue }
    if ($newPn -and -not (Test-PartNumberMatchesArenaNumberFormatFields -PartNumber $newPn -FieldDefs $fields)) {
      continue
    }
    $resolvedGuid = if ("$($full.guid)".Trim()) { "$($full.guid)".Trim() } else { $g }
    $name = if ("$($full.name)".Trim()) { "$($full.name)".Trim() } elseif ("$($row.name)".Trim()) { "$($row.name)".Trim() } else { $resolvedGuid }
    $score = Get-ArenaNumberFormatMatchSpecificityScore -PartNumber $pn -FieldDefs $fields
    $matches += ,([ordered]@{
        guid   = $resolvedGuid
        fields = $fields
        name   = $name
        score  = $score
      })
  }

  if ($matches.Count -eq 0) { return $null }
  if ($matches.Count -eq 1) {
    $m = $matches[0]
    return [ordered]@{ guid = $m.guid; fields = $m.fields }
  }

  $bestScore = ($matches | ForEach-Object { [int]$_.score } | Measure-Object -Maximum).Maximum
  $best = @($matches | Where-Object { [int]$_.score -eq $bestScore })
  if ($best.Count -gt 1) {
    $names = ($best | ForEach-Object { $_.name }) -join ', '
    throw (
      "Part number '$pn' matches more than one Arena number format ($names) for this duplicate. " +
      "Assign a number format on the item category in Arena settings, or adjust -NewPartNumber."
    )
  }
  $m = $best[0]
  return [ordered]@{ guid = $m.guid; fields = $m.fields }
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
    [Parameter(Mandatory)][string]$ApiRoot,
    [string]$NewPartNumber = ''
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

  $partNumber = ("$($SourceItem.number)").Trim()

  if (-not $nfGuid) {
    return Resolve-ArenaNumberFormatDefinitionFromPartNumber -PartNumber $partNumber -NewPartNumber $NewPartNumber -Headers $Headers -ApiRoot $ApiRoot
  }

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

  return Resolve-ArenaNumberFormatDefinitionFromPartNumber -PartNumber $partNumber -NewPartNumber $NewPartNumber -Headers $Headers -ApiRoot $ApiRoot
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

  $defs = Get-ArenaItemNumberFormatDefinition -SourceItem $SourceItem -Headers $Headers -ApiRoot $ApiRoot -NewPartNumber $newPn
  if (-not $defs -or -not $defs.guid) { return $null }
  $nfGuid = "$($defs.guid)".Trim()
  $fieldDefs = @($defs.fields)
  if ($fieldDefs.Count -eq 0) { return $null }

  $valFieldsOnly = Get-ArenaNumberFormatValueFields -FieldDefs $fieldDefs
  if (@($valFieldsOnly).Count -eq 1) {
    $t0 = ("$($valFieldsOnly[0].type)").Trim().ToUpperInvariant()
    if ($t0 -eq 'FREE_TEXT') {
      $gf = "$($valFieldsOnly[0].guid)".Trim()
      if (-not $gf) { return $null }
      return [ordered]@{ guid = $nfGuid; fields = @(@{ guid = $gf; value = $newPn }) }
    }
  }

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
      # Arena ignores values for AUTO_SEQUENCE on create and assigns the next
      # sequence itself. Omit those fields from the create body (same as Arena docs).
      if ($t -eq 'AUTO_SEQUENCE') {
        $segIdx++
        continue
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
    if ($outFields.Count -eq 0) { return $null }
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
$useProductImage = $UseProductImage -match '^(?i)yes|1|true|on$'
$customImagePath = if ($null -ne $ItemImageFilePath) { $ItemImageFilePath.Trim() } else { '' }
$passedImageGuid = if ($null -ne $ItemImageFileGuid) { $ItemImageFileGuid.Trim() } else { '' }
$doDisplayImage = [bool]$passedImageGuid -or $useProductImage -or [bool]$customImagePath
if (-not $passedImageGuid -and -not $useProductImage -and -not $customImagePath) {
  throw 'Item image: tick Use Product Image, or choose a local image file.'
}
if ($customImagePath -and -not $useProductImage -and -not $passedImageGuid -and -not (Test-Path -LiteralPath $customImagePath)) {
  throw "Item image file not found: $customImagePath"
}

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

$sourceProductImageAssoc = $null
if ($doDisplayImage -and $useProductImage -and -not $passedImageGuid) {
  try {
    $srcFilesForImage = Invoke-ArenaJson -Uri "$apiRoot/items/$srcRevGuid/files" -Headers $headers
  } catch {
    throw "Could not list files on source revision for Product Image: $($_.Exception.Message)"
  }
  $sourceProductImageAssoc = Select-ProductImageAssociation $srcFilesForImage
  if (-not $sourceProductImageAssoc) {
    throw "Use Product Image is on, but the source revision has no Illustrative Image / Product Image file."
  }
}

$numberFormatBody = Build-NumberFormatForCreateFromSourceItem -SourceItem $src -NewPartNumber $newPn -Headers $headers -ApiRoot $apiRoot
if (-not $numberFormatBody) {
  throw (
    "This workspace does not allow setting item 'number' directly on create (Arena 3031). " +
    "Could not resolve a number format: no usable numberFormat on GET /items or the items lookup, " +
    "no field list from GET /settings/items/categories (including parent categories) or GET /settings/items/numberformats, " +
    "and the source part number did not uniquely match any workspace number format. " +
    "Category guid: $($src.category.guid). Confirm this category (or an ancestor) has an item number format in Arena settings."
  )
}

$createDescription = $ItemDescription.Trim()
if (-not $createDescription) {
  $createDescription = $srcDesc
}
if (-not $createDescription) {
  $createDescription = $ItemName.Trim()
  if ($createDescription) {
    Write-Warning 'Item description was blank and the source item has no description; using item name for Arena create (description is required).'
  }
}
if (-not $createDescription) {
  throw 'Arena requires a non-empty item description. Enter text in Item description or use Load name & description from Arena.'
}

$create = [ordered]@{
  category       = @{ guid = "$($src.category.guid)" }
  name           = $ItemName.Trim()
  description    = $createDescription
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

$assignedNumber = if ($null -ne $created.number) { "$($created.number)".Trim() } else { '' }
if (-not $assignedNumber) {
  try {
    $createdFresh = Invoke-ArenaJson -Uri "$apiRoot/items/$($newGuid)?responseview=full" -Headers $headers
    if ($createdFresh -and $null -ne $createdFresh.number) {
      $assignedNumber = "$($createdFresh.number)".Trim()
    }
  } catch {}
}
Write-Host ("  Arena number:  {0}" -f $(if ($assignedNumber) { $assignedNumber } else { '(none)' }))
if ($assignedNumber -ne $newPn) {
  $renamed = Set-ArenaItemExactNumber -ItemGuid $newGuid -DesiredNumber $newPn -Headers $headers -ApiRoot $apiRoot
  $assignedNumber = if ($renamed -and $null -ne $renamed.number) { "$($renamed.number)".Trim() } else { '' }
  if (-not $assignedNumber -or $assignedNumber -ne $newPn) {
    # Confirm with GET in case PUT response omits number.
    $confirm = Invoke-ArenaJson -Uri "$apiRoot/items/$($newGuid)?responseview=full" -Headers $headers
    $assignedNumber = if ($confirm -and $null -ne $confirm.number) { "$($confirm.number)".Trim() } else { $assignedNumber }
  }
  if ($assignedNumber -ne $newPn) {
    throw (
      "Created item guid=$newGuid but could not set number to '$newPn' " +
      "(Arena has '$assignedNumber'). Open the item in Arena and correct the number, or delete it and retry."
    )
  }
  Write-Host ("  Confirmed number: {0}" -f $assignedNumber) -ForegroundColor DarkGreen
}

$newRevs = Invoke-ArenaJson -Uri "$apiRoot/items/$newGuid/revisions" -Headers $headers
$newWorking = $newRevs.results | Where-Object { $_.status -eq 0 } | Select-Object -First 1
if (-not $newWorking) { throw 'New item has no working revision (status 0).' }
$newWorkingRevGuid = "$($newWorking.guid)".Trim()

if ($doBom) {
  Write-Host '=== Copy BOM lines ==='
  try {
    $bom = Invoke-ArenaJson -Uri "$apiRoot/items/$srcGuid/bom" -Headers $headers
  } catch {
    throw "Could not read BOM on source item: $($_.Exception.Message)"
  }
  if ($bom -and $bom.results) {
    $n = 0
    $attempted = 0
    $failed = New-Object System.Collections.Generic.List[string]
    foreach ($line in @($bom.results)) {
      if (-not $line.item -or -not $line.item.guid) { continue }
      $attempted++
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
        $msg = "refDes=$($line.refDes): $($_.Exception.Message)"
        [void]$failed.Add($msg)
        Write-Warning "BOM line add failed ($msg)"
      }
    }
    Write-Host ("  Added {0} of {1} BOM line(s)." -f $n, $attempted)
    if ($attempted -gt 0 -and $n -eq 0) {
      throw ("BOM copy failed: none of the {0} source BOM line(s) could be added to the new item.{1}{2}" -f `
        $attempted, [Environment]::NewLine, ($failed -join [Environment]::NewLine))
    }
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
      if ($doDisplayImage -and (Test-IsIllustrativeImageCategory $a)) {
        Write-Host ("  Skipping Illustrative Image (handled as item image): {0}" -f $a.file.title)
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

$linkedImageGuid = ''
if ($doDisplayImage) {
  Write-Host '=== Item display image ==='
  if ($passedImageGuid) {
    $linkedImageGuid = $passedImageGuid
    Write-Host ("  Using shared file guid {0}" -f $linkedImageGuid)
  }
  elseif ($useProductImage) {
    $prodAssoc = $sourceProductImageAssoc
    if (-not $prodAssoc) {
      throw "Use Product Image is on, but the source revision has no Illustrative Image / Product Image file."
    }
    $linkedImageGuid = "$($prodAssoc.file.guid)".Trim()
    Write-Host ("  Source Product Image: {0} (guid {1})" -f $prodAssoc.file.title, $linkedImageGuid)
  }
  else {
    $revLetterOnly = Get-DuplicateRevLetterOnly $(if ($NewItemRevision.Trim()) { $NewItemRevision.Trim() } else { 'A' })
    $imageTitle = "Product Image $newPn Rev $revLetterOnly"
    $catGuid = Resolve-IllustrativeImageCategoryGuid -Headers $headers -ApiRoot $apiRoot
    $linkedImageGuid = New-ArenaFileFromLocalPath `
      -FilePath $customImagePath `
      -Title $imageTitle `
      -CategoryGuid $catGuid `
      -SessionId $sessionId `
      -ApiRoot $apiRoot `
      -Author $OwnerFullName
  }
  Add-ArenaRevisionFileAssociation -RevGuid $newWorkingRevGuid -FileGuid $linkedImageGuid -Headers $headers -ApiRoot $apiRoot
}

Write-Host ''
if ($BarePcbFollowUp) {
  Write-Host 'Done. Bare PCB duplicate finished.' -ForegroundColor DarkGreen
} else {
  Write-Host 'Done. Primary duplicate finished.' -ForegroundColor DarkGreen
}

function Convert-PcbaTextToPcbBareText {
  param([string]$Text)
  if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
  return ($Text -replace '(?i)PCBA', 'PCB')
}

$alsoBare = (-not $BarePcbFollowUp) -and ($AddBarePcbAlso -match '^(?i)yes|1|true|on$')
if ($alsoBare) {
  $helpers = Join-Path $scriptsDir 'AltiumRevisionHelpers.ps1'
  if (-not (Test-Path -LiteralPath $helpers)) {
    throw "Required script missing: $helpers"
  }
  . $helpers

  $pcbaOld = $OldPartNumber.Trim()
  $oldBareVariants = Get-BareBoardPartNumberVariants -PartNumber $pcbaOld
  $newBareVariants = Get-BareBoardPartNumberVariants -PartNumber $newPn
  $bareRev = Get-RevisionLetterOnly -RevisionText $(if ($NewItemRevision.Trim()) { $NewItemRevision.Trim() } else { 'A' })

  $bareOld = $null
  $bareNew = $null
  foreach ($pair in @(
      @{ Old = $oldBareVariants.Standard; New = $newBareVariants.Standard },
      @{ Old = $oldBareVariants.Legacy;   New = $newBareVariants.Legacy }
    )) {
    $pnBareEnc = [Uri]::EscapeDataString($pair.Old)
    $bareLookup = Invoke-ArenaJson -Uri "$apiRoot/items?number=$pnBareEnc&limit=1&responseview=full" -Headers $headers
    if ($bareLookup.results -and $bareLookup.results.Count -gt 0) {
      $bareOld = $pair.Old
      $bareNew = $pair.New
      break
    }
  }
  if (-not $bareOld) {
    throw (
      "Bare PCB duplicate: no Arena item found for bare board part number '$($oldBareVariants.Standard)' " +
      "or legacy '$($oldBareVariants.Legacy)' (mapped from PCBA '$pcbaOld')."
    )
  }

  $bareItemName = Convert-PcbaTextToPcbBareText $ItemName.Trim()
  $bareItemDesc = Convert-PcbaTextToPcbBareText $createDescription
  if (-not $bareItemName) {
    throw 'Bare PCB duplicate: PCBA item name is empty; cannot derive bare board name.'
  }

  Write-Host ''
  Write-Host '=== Bare PCB duplicate (auto from PCBA) ===' -ForegroundColor Cyan
  Write-Host ("  Source bare:  {0}" -f $bareOld)
  Write-Host ("  New bare:     {0}" -f $bareNew)
  Write-Host ("  Revision:     {0} (letter only)" -f $bareRev)
  Write-Host ("  Item name:    {0}" -f $bareItemName)
  if ($bareItemDesc -and $bareItemDesc -ne $bareItemName) {
    Write-Host ("  Description:  {0}" -f $bareItemDesc)
  }
  Write-Host ''

  # Hashtable splat binds by name. An array of '-Name','value' pairs is positional,
  # so the image guid is left over and the follow-up call fails to bind.
  $bareImageArgs = @{}
  if ($linkedImageGuid) {
    $bareImageArgs.UseProductImage = 'no'
    $bareImageArgs.ItemImageFileGuid = $linkedImageGuid
  }

  & $PSCommandPath `
    -OldPartNumber $bareOld `
    -NewPartNumber $bareNew `
    -ItemName $bareItemName `
    -ItemDescription $bareItemDesc `
    -CopyBom $CopyBom `
    -CopyFiles $CopyFiles `
    -SourceRevisionMode $SourceRevisionMode `
    -SourceRevisionOverride $SourceRevisionOverride `
    -NewItemRevision $bareRev `
    -OwnerFullName $OwnerFullName `
    -Mode Execute `
    -AddBarePcbAlso 'no' `
    -BarePcbFollowUp `
    -Email $Email `
    -WorkspaceId $WorkspaceId `
    -BaseUrl $BaseUrl `
    @bareImageArgs
  if ($LASTEXITCODE -ne 0) {
    throw "Bare PCB duplicate failed (exit $LASTEXITCODE)."
  }
}

Write-Host ''
Write-Host 'Done. Duplicate item flow finished.' -ForegroundColor DarkGreen
exit 0
