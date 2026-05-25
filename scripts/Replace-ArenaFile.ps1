<#
.SYNOPSIS
  Create a new edition of an existing Arena file with replacement content.

.DESCRIPTION
  POSTs multipart/form-data to /files/<FileGuid>/editions with new bytes,
  bumped edition number, and a new title taken from the replacement
  filename (without extension).

  Notes from Arena docs:
    - When uploading content, the response is empty; success is HTTP 201.
    - X-Arena-FileGuid in the response identifies the new edition.
    - The API may return 400 if the file is locked / checked out; this script still attempts
      the POST when the file reports locked so a new edition can be created when Arena allows it.

  Credentials via env: ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID.

.PARAMETER FileGuid          GUID of the existing Arena file (not the
                             association GUID).
.PARAMETER FilePath          Local replacement file.
.PARAMETER CurrentEdition    Current edition string from Arena (e.g. "3").
                             The new edition will be (parsed as int) + 1.
.PARAMETER NewEdition        Optional explicit override of the new edition.
.PARAMETER Title             Optional title override; default = replacement
                             filename without extension.
.PARAMETER Author            Author full name for the new edition. When omitted or blank,
                             uses the Windows interactive session account (same as upload).
#>
param(
  [Parameter(Mandatory = $true)] [string]$FileGuid,
  [Parameter(Mandatory = $true)] [string]$FilePath,
  [string]$CurrentEdition = '',
  [string]$NewEdition = '',
  [string]$Title = '',
  [string]$Author = '',
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
if ([string]::IsNullOrWhiteSpace($password))    { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($WorkspaceId)) { throw 'Set ARENA_WORKSPACE_ID or pass -WorkspaceId.' }
if (-not (Test-Path -LiteralPath $FilePath))    { throw "File not found: $FilePath" }

# Compute new edition if not given.
if ([string]::IsNullOrWhiteSpace($NewEdition)) {
  if ([string]::IsNullOrWhiteSpace($CurrentEdition)) {
    throw 'Provide -NewEdition or -CurrentEdition so the script can bump it.'
  }
  $cur = 0
  if (-not [int]::TryParse($CurrentEdition.Trim(), [ref]$cur)) {
    throw "Cannot parse current edition '$CurrentEdition' as an integer; pass -NewEdition explicitly."
  }
  $NewEdition = ($cur + 1).ToString()
}

$apiRoot = $BaseUrl.TrimEnd('/')

Add-Type -AssemblyName System.Net.Http -ErrorAction Stop

function Invoke-ArenaJson {
  param([string]$Uri, [string]$Method = 'Get', [hashtable]$Headers, [string]$Body, [string]$ContentType = 'application/json')
  $params = @{ Uri = $Uri; Method = $Method; UseBasicParsing = $true }
  if ($Headers) { $params.Headers = $Headers }
  if ($Body) { $params.Body = $Body; $params.ContentType = $ContentType }
  $resp = Invoke-WebRequest @params
  $bytes = $resp.RawContentStream.ToArray()
  $text = [System.Text.Encoding]::UTF8.GetString($bytes)
  if ([string]::IsNullOrWhiteSpace($text)) { return $null }
  return $text | ConvertFrom-Json
}

# ---- login ----

$loginBody = @{ email = $Email; password = $password; workspaceId = [long]$WorkspaceId } | ConvertTo-Json
try {
  $login = Invoke-ArenaJson -Uri "$apiRoot/login" -Method Post -Body $loginBody
} catch { throw "Arena login failed: $($_.Exception.Message)" }
$sid = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sid)) { $sid = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sid)) { throw 'Login response did not include arenaSessionId.' }
$jsonHeaders = @{ arena_session_id = $sid; 'Content-Type' = 'application/json' }

# ---- gather existing file metadata so we can show before/after info ----

$fileDetails = Invoke-ArenaJson -Uri "$apiRoot/files/$FileGuid" -Headers $jsonHeaders
if (-not $fileDetails) { throw "Could not fetch existing file details for $FileGuid." }

$fileInfo = Get-Item -LiteralPath $FilePath
$fileName = $fileInfo.Name
$ext = [System.IO.Path]::GetExtension($fileName).TrimStart('.')
$fileFormat = if ([string]::IsNullOrWhiteSpace($ext)) { 'bin' } else { $ext }

if ([string]::IsNullOrWhiteSpace($Title)) {
  $Title = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
}

Write-Output ("Replacing file:  GUID {0}" -f $FileGuid)
Write-Output ("  current edition: {0}  ->  new edition: {1}" -f $fileDetails.edition, $NewEdition)
Write-Output ("  current title:   '{0}'  ->  new title: '{1}'" -f $fileDetails.title, $Title)
Write-Output ("  category (kept): {0}" -f $fileDetails.category.path)
Write-Output ("  new content:     {0} ({1} bytes)" -f $fileName, $fileInfo.Length)
Write-Output ("  session author:  {0} (not sent in multipart; Arena uses login user)" -f $Author)
if ($fileDetails.locked) {
  Write-Host (
    "Note: Arena GET /files reports locked=true for $($fileDetails.number); " +
    "posting edition $NewEdition anyway (matches manual replace in the UI when not checked out)."
  ) -ForegroundColor DarkGray
}

# ---- hand-built multipart body ----
#
# Surprise: POST /files/<guid>/editions accepts fields ONLY under the
# `file.` dotted prefix (e.g. file.edition, file.title). Sending them
# as plain `edition`, `title`, etc. returns 400 "attribute X is not
# recognized" even though the docs sample shows them unprefixed.
#
# Category/format/etc. are inherited from the existing file. Do not send
# author.fullName on this endpoint — Arena returns 4074 ("attribute not recognized")
# when author is sent unprefixed; file.author.fullName is also rejected on some
# workspaces. The logged-in Arena user is credited automatically.

$boundary = [System.Guid]::NewGuid().ToString('N')
$enc = [System.Text.Encoding]::UTF8
$LF = "`r`n"
$bodyStream = New-Object System.IO.MemoryStream

function Write-StringToStream { param([System.IO.MemoryStream]$Stream, [string]$Text)
  $b = $enc.GetBytes($Text); $Stream.Write($b, 0, $b.Length)
}
function Add-MultipartField { param([System.IO.MemoryStream]$Stream, [string]$Boundary, [string]$Name, [string]$Value)
  if ($null -eq $Value) { $Value = '' }
  Write-StringToStream -Stream $Stream -Text ("--$Boundary$LF" +
    "Content-Disposition: form-data; name=`"$Name`"$LF$LF" +
    "$Value$LF")
}
function Add-MultipartFile { param([System.IO.MemoryStream]$Stream, [string]$Boundary, [string]$Name, [string]$FileName, [byte[]]$Bytes)
  Write-StringToStream -Stream $Stream -Text ("--$Boundary$LF" +
    "Content-Disposition: form-data; name=`"$Name`"; filename=`"$FileName`"$LF" +
    "Content-Type: application/octet-stream$LF$LF")
  $Stream.Write($Bytes, 0, $Bytes.Length)
  Write-StringToStream -Stream $Stream -Text $LF
}

Add-MultipartField $bodyStream $boundary 'file.edition' $NewEdition
Add-MultipartField $bodyStream $boundary 'file.title'   $Title
Add-MultipartField $bodyStream $boundary 'file.format'  $fileFormat

$fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
Add-MultipartFile $bodyStream $boundary 'content' $fileName $fileBytes
Write-StringToStream -Stream $bodyStream -Text ("--$boundary--$LF")

$bodyBytes = $bodyStream.ToArray()

$client = New-Object System.Net.Http.HttpClient
$client.Timeout = [TimeSpan]::FromMinutes(5)
$client.DefaultRequestHeaders.Add('arena_session_id', $sid)

$reqContent = New-Object System.Net.Http.ByteArrayContent (,$bodyBytes)
$reqContent.Headers.ContentType =
  [System.Net.Http.Headers.MediaTypeHeaderValue]::Parse("multipart/form-data; boundary=$boundary")

$uri = "$apiRoot/files/$FileGuid/editions"
try {
  $resp = $client.PostAsync($uri, $reqContent).GetAwaiter().GetResult()
} catch {
  throw "Arena POST /files/$FileGuid/editions transport error: $($_.Exception.Message)"
}

$respBody = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
if (-not $resp.IsSuccessStatusCode) {
  $code = [int]$resp.StatusCode
  $detail = if ([string]::IsNullOrWhiteSpace($respBody)) { '<empty body>' } else { $respBody }
  throw ("Arena POST /files/$FileGuid/editions returned {0}: {1}" -f $code, $detail)
}

$newGuid = $null
if ($resp.Headers.Contains('X-Arena-FileGuid')) {
  $newGuid = ($resp.Headers.GetValues('X-Arena-FileGuid') | Select-Object -First 1)
}

$displayGuid = if ([string]::IsNullOrWhiteSpace($newGuid)) { '<not returned>' } else { $newGuid }
Write-Output ("New edition created. HTTP {0}.  X-Arena-FileGuid: {1}" -f [int]$resp.StatusCode, $displayGuid)
Write-Output "Done. View in Arena: https://app.bom.com/$FileGuid"

exit 0
