<#
.SYNOPSIS
  Find an Arena item by manufacturer part number and open it in the Arena web app.

.DESCRIPTION
  Searches supplier items whose number contains the lookup text (partial, case-insensitive),
  then follows each supplier item's sourcing links to Arena internal items. When exactly one
  internal part number is found, prints it and opens the item page in your default browser.

.PARAMETER ManufacturerPartNumber
  Manufacturer part number or partial MPN (e.g. S3D-E3/57 or ABC123).

.EXAMPLE
  .\Open-PartInArena.ps1 -ManufacturerPartNumber S3D-E3/57T
#>
param(
  [Parameter(Mandatory = $true)] [string]$ManufacturerPartNumber,
  [string]$Email = $env:ARENA_EMAIL,
  [string]$WorkspaceId = $env:ARENA_WORKSPACE_ID,
  [string]$BaseUrl = 'https://api.arenasolutions.com/v1'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

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

function Normalize-PartKey([string]$Value) {
  if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
  return $Value.Trim().ToLowerInvariant()
}

function Test-PartNumberPartialMatch {
  param(
    [string]$Candidate,
    [string]$Needle
  )
  $haystack = Normalize-PartKey $Candidate
  if (-not $haystack -or -not $Needle) { return $false }
  return $haystack.Contains($Needle)
}

function Find-ArenaItemsByManufacturerPartNumber {
  param(
    [string]$Lookup,
    [string]$ApiRoot,
    [hashtable]$Headers
  )
  $needle = Normalize-PartKey $Lookup
  $pattern = "*$Lookup*"
  $patternEnc = [Uri]::EscapeDataString($pattern)
  $resp = Invoke-ArenaJson -Uri "$ApiRoot/supplieritems?number=$patternEnc&limit=100" -Headers $Headers
  if (-not $resp -or -not $resp.results) { return @() }

  $byNumber = @{}
  foreach ($si in @($resp.results)) {
    if (-not (Test-PartNumberPartialMatch -Candidate $si.number -Needle $needle)) { continue }

    $src = Invoke-ArenaJson -Uri "$ApiRoot/supplieritems/$($si.guid)/sourcing" -Headers $Headers
    if (-not $src -or -not $src.results) { continue }

    foreach ($rel in @($src.results)) {
      if (-not $rel.item -or -not $rel.item.number) { continue }
      $num = ("$($rel.item.number)").Trim()
      if (-not $num) { continue }

      if (-not $byNumber.ContainsKey($num)) {
        $byNumber[$num] = [ordered]@{
          Number              = $num
          Name                = if ($rel.item.name) { ("$($rel.item.name)").Trim() } else { '' }
          MatchedSupplierParts = New-Object System.Collections.Generic.List[string]
        }
      }
      $siNum = ("$($si.number)").Trim()
      if ($siNum -and ($byNumber[$num].MatchedSupplierParts -notcontains $siNum)) {
        [void]$byNumber[$num].MatchedSupplierParts.Add($siNum)
      }
    }
  }

  return @($byNumber.Values | Sort-Object { $_.Number })
}

function Resolve-ArenaItemAppUrl {
  param(
    [string]$ItemNumber,
    [string]$ApiRoot,
    [hashtable]$Headers
  )
  $pnEnc = [Uri]::EscapeDataString($ItemNumber.Trim())
  $lookup = Invoke-ArenaJson -Uri "$ApiRoot/items?number=$pnEnc&limit=1" -Headers $Headers
  if (-not $lookup.results -or $lookup.results.Count -eq 0) { return $null }
  $item = $lookup.results[0]
  if ($item.url -and $item.url.app) { return ("$($item.url.app)").Trim() }
  return $null
}

$password = $env:ARENA_PASSWORD
if ([string]::IsNullOrWhiteSpace($Email)) { throw 'Set ARENA_EMAIL or pass -Email.' }
if ([string]::IsNullOrWhiteSpace($password)) { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($WorkspaceId)) { throw 'Set ARENA_WORKSPACE_ID or pass -WorkspaceId.' }

$lookupRaw = $ManufacturerPartNumber.Trim()
if (-not $lookupRaw) { throw 'Manufacturer part number is empty.' }

$apiRoot = $BaseUrl.TrimEnd('/')

$loginBody = @{ email = $Email; password = $password; workspaceId = [long]$WorkspaceId } | ConvertTo-Json
try {
  $login = Invoke-ArenaJson -Uri "$apiRoot/login" -Method Post -Body $loginBody
} catch {
  throw "Arena login failed: $($_.Exception.Message)"
}

$sid = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sid)) { $sid = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sid)) { throw 'Login response did not include arenaSessionId.' }

$headers = @{ arena_session_id = $sid; 'Content-Type' = 'application/json' }

Write-Output '=== Open Part in Arena ==='
Write-Output ("  MPN lookup:     {0}" -f $lookupRaw)
Write-Output ''

$matches = @(Find-ArenaItemsByManufacturerPartNumber -Lookup $lookupRaw -ApiRoot $apiRoot -Headers $headers)

if ($matches.Count -eq 0) {
  throw (
    "No Arena item linked to manufacturer part number '$lookupRaw' " +
    '(searched supplier items; partial match).'
  )
}

if ($matches.Count -gt 1) {
  Write-Output ("Found {0} Arena part numbers; not opening a page:" -f $matches.Count)
  foreach ($m in $matches) {
    $mpnList = ($m.MatchedSupplierParts -join ', ')
    Write-Output ("  - {0}  {1}  (via {2})" -f $m.Number, $m.Name, $mpnList)
  }
  exit 2
}

$hit = $matches[0]
$mpnMatched = ($hit.MatchedSupplierParts -join ', ')
Write-Output ("Arena part number: {0}" -f $hit.Number)
if ($hit.Name) { Write-Output ("Name:              {0}" -f $hit.Name) }
if ($mpnMatched) { Write-Output ("Matched MPN:       {0}" -f $mpnMatched) }

$appUrl = Resolve-ArenaItemAppUrl -ItemNumber $hit.Number -ApiRoot $apiRoot -Headers $headers
if (-not $appUrl) {
  throw "Could not resolve Arena app URL for part number '$($hit.Number)'."
}

Write-Output ("Opening:           {0}" -f $appUrl)
Start-Process -FilePath $appUrl
Write-Output 'Done.'
exit 0
