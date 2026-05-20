<#
.SYNOPSIS
  Upload a file to an Arena PLM item's working revision (or a chosen revision).

.DESCRIPTION
  Three-step Arena workflow:
    1. Look up the item by internal part number to get an item GUID.
    2. GET /items/<guid>/revisions to find a target revision GUID:
         - default Revision='working'  → status=0
         - Revision='effective'        → status=1
         - Revision='<rev letter>'     → matches that revision number
    3. Resolve the file Category by name or path (e.g. 'Test Specification' or
       'File\Supplier File\Test Specification') to a category GUID.
    4. POST /files (multipart/form-data) with the file content + metadata.
       The new file GUID comes back in response header X-Arena-FileGuid.
    5. POST /items/<targetRevGuid>/files (application/json) to associate the
       new file with that item revision.

  Credentials via env: ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID.

.PARAMETER PartNumber       Item internal part number, e.g. 294-10-001.
.PARAMETER FilePath         Local file to upload.
.PARAMETER Title            File title (shown in Arena).
.PARAMETER Description      File description (optional).
.PARAMETER Edition          File edition string (default '1').
.PARAMETER CategoryName     Category name OR full path. Default 'Test Specification'.
.PARAMETER Revision         'working' (default), 'effective', or a rev string like 'A','C1'.
.PARAMETER Author           Author full name. When omitted or blank, uses the Windows
                            interactive session account (display name when available,
                            else DOMAIN\User).
.PARAMETER FileFormat       Format/extension override (default = file extension).
.PARAMETER Private          Mark file private (default $false).

.EXAMPLE
  .\Upload-ArenaFile.ps1 -PartNumber 294-10-001 `
      -FilePath 'C:\fab-test-results.pdf' `
      -Title 'Fab Test Results 294-10-001' `
      -CategoryName 'Test Specification'
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [Parameter(Mandatory = $true)] [string]$FilePath,
  [Parameter(Mandatory = $true)] [string]$Title,
  [string]$Description = '',
  [string]$Edition = '1',
  [string]$CategoryName = 'Test Specification',
  [string]$Revision = 'working',
  [string]$Author = '',
  [string]$FileFormat = '',
  [switch]$Private,
  [string]$Email = $env:ARENA_EMAIL,
  [string]$WorkspaceId = $env:ARENA_WORKSPACE_ID,
  [string]$BaseUrl = 'https://api.arenasolutions.com/v1'
)

$ErrorActionPreference = 'Stop'

$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

. (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) 'ArenaSessionAuthor.ps1')
if ([string]::IsNullOrWhiteSpace($Author)) {
  $Author = Get-ArenaWindowsSessionAuthorName
}

$password = $env:ARENA_PASSWORD
if ([string]::IsNullOrWhiteSpace($Email))       { throw 'Set ARENA_EMAIL or pass -Email.' }
if ([string]::IsNullOrWhiteSpace($password))    { throw 'Set ARENA_PASSWORD (required for login).' }
if ([string]::IsNullOrWhiteSpace($WorkspaceId)) { throw 'Set ARENA_WORKSPACE_ID or pass -WorkspaceId.' }
if (-not (Test-Path -LiteralPath $FilePath))    { throw "File not found: $FilePath" }

$apiRoot = $BaseUrl.TrimEnd('/')

Add-Type -AssemblyName System.Net.Http -ErrorAction Stop

# ---------- shared helpers ----------

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

# ---------- login ----------

$loginBody = @{
  email       = $Email
  password    = $password
  workspaceId = [long]$WorkspaceId
} | ConvertTo-Json

try {
  $login = Invoke-ArenaJson -Uri "$apiRoot/login" -Method Post -Body $loginBody
} catch {
  throw "Arena login failed: $($_.Exception.Message)"
}

$sid = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sid)) { $sid = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sid)) { throw 'Login response did not include arenaSessionId.' }

$jsonHeaders = @{
  arena_session_id = $sid
  'Content-Type'   = 'application/json'
}

# ---------- look up item ----------

$pnEnc = [Uri]::EscapeDataString($PartNumber.Trim())
$lookup = Invoke-ArenaJson -Uri "$apiRoot/items?number=$pnEnc&limit=1" -Headers $jsonHeaders
if (-not $lookup.results -or $lookup.results.Count -eq 0) {
  throw "No item found for number '$($PartNumber.Trim())'."
}
$itemAnyGuid = $lookup.results[0].guid

$revs = Invoke-ArenaJson -Uri "$apiRoot/items/$itemAnyGuid/revisions" -Headers $jsonHeaders
if (-not $revs.results -or $revs.results.Count -eq 0) {
  throw "No revisions returned for item $($PartNumber.Trim())."
}

$revKey = $Revision.Trim().ToLowerInvariant()
$targetRev = switch ($revKey) {
  'working'   { $revs.results | Where-Object { $_.status -eq 0 } | Select-Object -First 1 }
  'effective' { $revs.results | Where-Object { $_.status -eq 1 } | Select-Object -First 1 }
  default     {
    $match = $revs.results |
      Where-Object { ($_.number -ne $null) -and ($_.number.ToString().ToLowerInvariant() -eq $revKey) } |
      Select-Object -First 1
    if (-not $match) { throw "No revision matches '$Revision'. Available: $((@($revs.results | ForEach-Object { '[' + ($_.status) + ' ' + $_.number + ']' })) -join ', ')" }
    $match
  }
}
if (-not $targetRev) {
  throw "Could not find revision '$Revision' for item $($PartNumber.Trim()). Try -Revision effective or a specific rev letter."
}
$targetRevGuid = $targetRev.guid
Write-Output ("Target revision: status={0} number={1} guid={2}" -f $targetRev.status, $targetRev.number, $targetRevGuid)
Write-Output ("Author:          {0}" -f $Author)

# ---------- resolve file category ----------

$cats = Invoke-ArenaJson -Uri "$apiRoot/settings/files/categories" -Headers $jsonHeaders
if (-not $cats.results) { throw 'Could not fetch file categories.' }
$catKey = $CategoryName.Trim().ToLowerInvariant()
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
  $sample = ($cats.results | Where-Object { $_.assignable -eq $true } | Select-Object -ExpandProperty path | Sort-Object) -join "`n  "
  throw "File category not found for '$CategoryName'. Assignable categories:`n  $sample"
}
$categoryGuid = $matchedCat.guid
Write-Output ("Category:        {0}  ({1})" -f $matchedCat.path, $categoryGuid)

# ---------- multipart upload (hand-built body) ----------
#
# Arena's /files endpoint is picky: it rejects the boundary-with-quotes that
# .NET's MultipartFormDataContent emits, and rejects per-part Content-Type
# headers on simple text fields (returns "storageMethodName not valid",
# "file content is missing"). So we build the body ourselves and POST raw bytes.

$fileInfo = Get-Item -LiteralPath $FilePath
$fileName = $fileInfo.Name
$ext = [System.IO.Path]::GetExtension($fileName).TrimStart('.')
if ([string]::IsNullOrWhiteSpace($FileFormat)) { $FileFormat = $ext }
if ([string]::IsNullOrWhiteSpace($FileFormat)) { $FileFormat = 'bin' }
$privateStr = if ($Private) { 'true' } else { 'false' }

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
Add-MultipartField $bodyStream $boundary 'category.guid'     $categoryGuid
Add-MultipartField $bodyStream $boundary 'title'             $Title
Add-MultipartField $bodyStream $boundary 'edition'           $Edition
Add-MultipartField $bodyStream $boundary 'format'            $FileFormat
Add-MultipartField $bodyStream $boundary 'private'           $privateStr
# Arena rejects blank values for optional string attributes (e.g. error 3215
# for description), so omit the field entirely when the user left it empty.
if (-not [string]::IsNullOrWhiteSpace($Description)) {
  Add-MultipartField $bodyStream $boundary 'description' $Description
}
Add-MultipartField $bodyStream $boundary 'author.fullName' $Author

$fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
Add-MultipartFile $bodyStream $boundary 'content' $fileName $fileBytes

Write-StringToStream -Stream $bodyStream -Text ("--$boundary--$LF")

$bodyBytes = $bodyStream.ToArray()
$createUri = "$apiRoot/files"

Write-Output ("Uploading:       {0} ({1} bytes), format='{2}', edition='{3}', private={4}" -f `
  $fileName, $fileInfo.Length, $FileFormat, $Edition, $privateStr)

# Use HttpClient (not Invoke-WebRequest) so we can read the response body even
# on a non-2xx — PS 5.1's Invoke-WebRequest drains the error stream first and
# leaves us with nothing useful to surface.
Add-Type -AssemblyName System.Net.Http -ErrorAction Stop
$httpClient = New-Object System.Net.Http.HttpClient
$httpClient.Timeout = [TimeSpan]::FromMinutes(5)
$httpClient.DefaultRequestHeaders.Add('arena_session_id', $sid)

$reqContent = New-Object System.Net.Http.ByteArrayContent (,$bodyBytes)
$reqContent.Headers.ContentType =
  [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("multipart/form-data; boundary=$boundary")

$resp = $null
try {
  $resp = $httpClient.PostAsync($createUri, $reqContent).GetAwaiter().GetResult()
} catch {
  throw "Arena POST /files transport error: $($_.Exception.Message)"
}

$respBody = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
if (-not $resp.IsSuccessStatusCode) {
  $code = [int]$resp.StatusCode
  $detail = if ([string]::IsNullOrWhiteSpace($respBody)) { '<empty body>' } else { $respBody }
  throw ("Arena POST /files returned {0}: {1}`n  request was: file='{2}' ({3} bytes), format='{4}', edition='{5}', title='{6}', category.guid='{7}'" -f `
    $code, $detail, $fileName, $fileInfo.Length, $FileFormat, $Edition, $Title, $categoryGuid)
}

$fileGuid = $null
if ($resp.Headers.Contains('X-Arena-FileGuid')) {
  $fileGuid = ($resp.Headers.GetValues('X-Arena-FileGuid') | Select-Object -First 1)
}
if ([string]::IsNullOrWhiteSpace($fileGuid) -and -not [string]::IsNullOrWhiteSpace($respBody)) {
  try { $fileGuid = ($respBody | ConvertFrom-Json).guid } catch {}
}
if ([string]::IsNullOrWhiteSpace($fileGuid)) {
  throw "File upload returned $([int]$resp.StatusCode) but no file GUID was found. Body: $respBody"
}
Write-Output ("Uploaded:        {0} ({1} bytes) -> fileGuid {2}" -f $fileName, $fileInfo.Length, $fileGuid)

# ---------- associate file with item revision ----------

$assocBody = @{
  latestEditionAssociation = $true
  primary                  = $false
  file                     = @{ guid = $fileGuid }
} | ConvertTo-Json

try {
  $assoc = Invoke-ArenaJson -Uri "$apiRoot/items/$targetRevGuid/files" `
    -Method Post -Headers $jsonHeaders -Body $assocBody
} catch {
  throw "Arena POST /items/$targetRevGuid/files failed: $($_.Exception.Message)"
}

Write-Output ("Associated:      file '{0}' (number {1}) -> item rev {2}" -f $assoc.file.title, $assoc.file.number, $targetRevGuid)
Write-Output ''
Write-Output "Done. View in Arena: https://app.bom.com/$targetRevGuid"

exit 0
