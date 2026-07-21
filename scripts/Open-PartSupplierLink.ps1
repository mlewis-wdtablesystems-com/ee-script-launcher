<#
.SYNOPSIS
  Find an Arena item by part number and open the primary supplier link.

.DESCRIPTION
  Searches the whole Arena workspace for a unique item match on internal part number
  (items search) or supplier / manufacturer part number (supplier items and sourcing).
  Lookup is a case-insensitive partial match (e.g. ABC123 matches ABC12345).

  When exactly one item matches, opens the primary file on the first sourcing location
  (lowest amlRank) for that item. WEB and FTP links open in the default browser.
  FILE associations download to a temp file and open with the default application.

.PARAMETER LookupPartNumber
  Internal Arena part number, supplier part number, or manufacturer part number to find.

.EXAMPLE
  .\Open-PartSupplierLink.ps1 -LookupPartNumber S3D-E3/57T
#>
param(
  [Parameter(Mandatory = $true)] [string]$LookupPartNumber,
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

function Invoke-ArenaBytes {
  param(
    [Parameter(Mandatory = $true)] [string]$Uri,
    [hashtable]$Headers
  )
  $resp = Invoke-WebRequest -Uri $Uri -Method Get -Headers $Headers -UseBasicParsing
  return $resp.RawContentStream.ToArray()
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

function Add-ArenaItemMatch {
  param(
    [hashtable]$Table,
    [string]$Number,
    [string]$Name,
    [string]$Guid,
    [string]$MatchedVia
  )
  $num = $Number.Trim()
  if (-not $num) { return }
  if (-not $Table.ContainsKey($num)) {
    $Table[$num] = [ordered]@{
      Number    = $num
      Name      = $Name.Trim()
      Guid      = $Guid.Trim()
      MatchedVia = New-Object System.Collections.Generic.List[string]
    }
  }
  if ($MatchedVia -and ($Table[$num].MatchedVia -notcontains $MatchedVia)) {
    [void]$Table[$num].MatchedVia.Add($MatchedVia)
  }
  if (-not $Table[$num].Guid -and $Guid) { $Table[$num].Guid = $Guid.Trim() }
  if (-not $Table[$num].Name -and $Name) { $Table[$num].Name = $Name.Trim() }
}

function Find-ArenaItemsByLookup {
  param(
    [string]$Lookup,
    [string]$ApiRoot,
    [hashtable]$Headers
  )
  $needle = Normalize-PartKey $Lookup
  $pattern = "*$Lookup*"
  $patternEnc = [Uri]::EscapeDataString($pattern)
  $byNumber = @{}

  $itemsResp = Invoke-ArenaJson -Uri "$ApiRoot/items?number=$patternEnc&limit=100" -Headers $Headers
  foreach ($it in @($itemsResp.results)) {
    if (-not $it -or -not $it.number) { continue }
    if (-not (Test-PartNumberPartialMatch -Candidate $it.number -Needle $needle)) { continue }
    Add-ArenaItemMatch -Table $byNumber -Number $it.number -Name $it.name -Guid $it.guid -MatchedVia 'item number'
  }

  $siResp = Invoke-ArenaJson -Uri "$ApiRoot/supplieritems?number=$patternEnc&limit=100" -Headers $Headers
  foreach ($si in @($siResp.results)) {
    if (-not $si -or -not $si.number) { continue }
    if (-not (Test-PartNumberPartialMatch -Candidate $si.number -Needle $needle)) { continue }

    $src = Invoke-ArenaJson -Uri "$ApiRoot/supplieritems/$($si.guid)/sourcing" -Headers $Headers
    if (-not $src -or -not $src.results) { continue }

    foreach ($rel in @($src.results)) {
      if (-not $rel.item -or -not $rel.item.number) { continue }
      Add-ArenaItemMatch -Table $byNumber -Number $rel.item.number -Name $rel.item.name `
        -Guid $rel.item.guid -MatchedVia ("supplier item $($si.number)")
    }
  }

  return @($byNumber.Values | Sort-Object { $_.Number })
}

function Resolve-ArenaItemGuid {
  param(
    [string]$ItemNumber,
    [string]$PreferredGuid,
    [string]$ApiRoot,
    [hashtable]$Headers
  )
  if ($PreferredGuid) { return $PreferredGuid.Trim() }
  $pnEnc = [Uri]::EscapeDataString($ItemNumber.Trim())
  $lookup = Invoke-ArenaJson -Uri "$ApiRoot/items?number=$pnEnc&limit=1" -Headers $Headers
  if (-not $lookup.results -or $lookup.results.Count -eq 0) { return $null }
  return ("$($lookup.results[0].guid)").Trim()
}

function Get-FirstSourcingRelationship {
  param(
    [string]$ItemGuid,
    [string]$ApiRoot,
    [hashtable]$Headers
  )
  $src = Invoke-ArenaJson -Uri "$ApiRoot/items/$ItemGuid/sourcing" -Headers $Headers
  if (-not $src -or -not $src.results -or $src.results.Count -eq 0) {
    return $null
  }
  return @($src.results | Sort-Object { [int]$_.amlRank }, { "$($_.guid)" }) | Select-Object -First 1
}

function Get-PrimarySupplierItemFileAssociation {
  param(
    [string]$SupplierItemGuid,
    [string]$ApiRoot,
    [hashtable]$Headers
  )
  $resp = Invoke-ArenaJson -Uri "$ApiRoot/supplieritems/$SupplierItemGuid/files" -Headers $Headers
  if (-not $resp -or -not $resp.results -or $resp.results.Count -eq 0) {
    return $null
  }
  $primary = @($resp.results | Where-Object { $_.primary -eq $true })
  if ($primary.Count -gt 0) { return $primary[0] }
  return $resp.results[0]
}

function Open-SupplierItemPrimaryFile {
  param(
    [object]$FileAssociation,
    [string]$ApiRoot,
    [hashtable]$Headers
  )
  if (-not $FileAssociation -or -not $FileAssociation.file) {
    throw 'No file association to open.'
  }

  $file = $FileAssociation.file
  $storage = ("$($file.storageMethodName)").Trim().ToUpperInvariant()
  $location = if ($null -ne $file.location) { ("$($file.location)").Trim() } else { '' }
  $title = if ($null -ne $file.title) { ("$($file.title)").Trim() } else { '' }
  $fileGuid = ("$($file.guid)").Trim()

  switch ($storage) {
    'WEB' {
      if (-not $location) { throw "Primary file '$title' is WEB but has no location URL." }
      Write-Output ("Opening web link: {0}" -f $location)
      Start-Process -FilePath $location
      return
    }
    'FTP' {
      if (-not $location) { throw "Primary file '$title' is FTP but has no location URL." }
      Write-Output ("Opening FTP link: {0}" -f $location)
      Start-Process -FilePath $location
      return
    }
    'FILE' {
      if (-not $fileGuid) { throw "Primary file '$title' is FILE but has no file GUID." }
      $bytes = Invoke-ArenaBytes -Uri "$ApiRoot/files/$fileGuid/content" -Headers $Headers
      if (-not $bytes -or $bytes.Length -eq 0) {
        throw "Primary file '$title' has no downloadable content."
      }
      $baseName = if ($file.name) { $file.name } elseif ($title) { $title } else { "supplier-file-$fileGuid" }
      foreach ($c in [System.IO.Path]::GetInvalidFileNameChars()) {
        $baseName = $baseName.Replace([string]$c, '_')
      }
      $ext = [System.IO.Path]::GetExtension($baseName)
      if (-not $ext -and $file.format) {
        $fmt = ("$($file.format)").Trim().ToLowerInvariant()
        if ($fmt -and $fmt -ne 'web link') {
          $ext = if ($fmt.StartsWith('.')) { $fmt } else { ".$fmt" }
          $baseName = "$baseName$ext"
        }
      }
      $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) $baseName
      if (Test-Path -LiteralPath $tempPath) {
        $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ("{0}-{1}{2}" -f [System.IO.Path]::GetFileNameWithoutExtension($baseName), [Guid]::NewGuid().ToString('N').Substring(0, 8), $ext)
      }
      [System.IO.File]::WriteAllBytes($tempPath, $bytes)
      Write-Output ("Opening downloaded file: {0}" -f $tempPath)
      Start-Process -FilePath $tempPath
      return
    }
    default {
      if ($location) {
        Write-Output ("Opening location for storageMethod=$storage : {0}" -f $location)
        Start-Process -FilePath $location
        return
      }
      throw "Primary file '$title' uses unsupported storage method '$storage' with no location."
    }
  }
}

$password = $env:ARENA_PASSWORD
if ([string]::IsNullOrWhiteSpace($Email)) { throw 'Set ARENA_EMAIL or pass -Email.' }
if ([string]::IsNullOrWhiteSpace($password)) { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($WorkspaceId)) { throw 'Set ARENA_WORKSPACE_ID or pass -WorkspaceId.' }

$lookupRaw = $LookupPartNumber.Trim()
if (-not $lookupRaw) { throw 'Lookup part number is empty.' }

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

Write-Output '=== Open Part Supplier Link ==='
Write-Output ("  Lookup part:    {0}" -f $lookupRaw)
Write-Output ''

$matches = @(Find-ArenaItemsByLookup -Lookup $lookupRaw -ApiRoot $apiRoot -Headers $headers)

if ($matches.Count -eq 0) {
  throw (
    "No Arena item matched lookup part '$lookupRaw' in the workspace " +
    '(internal, supplier, or manufacturer part number; partial match).'
  )
}

if ($matches.Count -gt 1) {
  Write-Output ("Found {0} matching Arena items; not opening a link:" -f $matches.Count)
  foreach ($m in $matches) {
    $via = ($m.MatchedVia -join '; ')
    Write-Output ("  - {0}  {1}  (via {2})" -f $m.Number, $m.Name, $via)
  }
  exit 2
}

$matched = $matches[0]
Write-Output ("Matched item: {0}  {1}" -f $matched.Number, $matched.Name)
if ($matched.MatchedVia.Count -gt 0) {
  Write-Output ("  via: {0}" -f ($matched.MatchedVia -join '; '))
}

$itemGuid = Resolve-ArenaItemGuid -ItemNumber $matched.Number -PreferredGuid $matched.Guid `
  -ApiRoot $apiRoot -Headers $headers
if (-not $itemGuid) {
  throw "Could not resolve Arena item GUID for '$($matched.Number)'."
}

$sourcing = Get-FirstSourcingRelationship -ItemGuid $itemGuid -ApiRoot $apiRoot -Headers $headers
if (-not $sourcing) {
  throw "Matched item '$($matched.Number)' has no sourcing locations."
}

$supplierItemGuid = ("$($sourcing.vendorItem.guid)").Trim()
$supplierName = if ($sourcing.vendorItem.supplier) { "$($sourcing.vendorItem.supplier.name)".Trim() } else { '' }
$vendorNumber = ("$($sourcing.vendorItem.number)").Trim()
if (-not $supplierItemGuid) {
  throw "First sourcing location on '$($matched.Number)' has no vendor supplier item."
}

Write-Output (
  "First sourcing location: rank={0} supplier={1} vendor part={2}" -f
  $sourcing.amlRank, $(if ($supplierName) { $supplierName } else { '(unknown)' }), $(if ($vendorNumber) { $vendorNumber } else { '(none)' })
)

$fileAssoc = Get-PrimarySupplierItemFileAssociation -SupplierItemGuid $supplierItemGuid -ApiRoot $apiRoot -Headers $headers
if (-not $fileAssoc) {
  throw (
    "No files on supplier item for vendor part '$vendorNumber' " +
    "$(if ($supplierName) { "($supplierName)" } else { '' })."
  )
}

$primaryTitle = if ($fileAssoc.file.title) { $fileAssoc.file.title } else { '(untitled)' }
Write-Output ("Primary file: {0}" -f $primaryTitle)

Open-SupplierItemPrimaryFile -FileAssociation $fileAssoc -ApiRoot $apiRoot -Headers $headers
Write-Output 'Done.'
exit 0
