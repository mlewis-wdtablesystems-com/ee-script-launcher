<#
.SYNOPSIS
  Recursively walk an Arena BOM and list every PCB assembly used in the
  system at any depth.

.DESCRIPTION
  A "PCB assembly" here means an Arena part number that matches the
  pattern PPP-10-YYY or PPP-15-YYY where PPP is NOT 098.

  Starts from -PartNumber, fetches the BOM via /items/<guid>/bom for
  every reachable child, and reports each unique PCBA found.

  Cycles / repeated sub-assemblies are visited only once (tracked by
  item GUID), but the count of how many BOM lines reference each PCBA
  and the path(s) where it was first seen are reported.

  Credentials via env: ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID.

.PARAMETER PartNumber
  Top-level system part number to walk from.

.PARAMETER Format
  table (default), json, or csv.

.PARAMETER MaxDepth
  Safety cap on recursion depth (default 20).

.EXAMPLE
  .\Find-ArenaPcbAssemblies.ps1 -PartNumber 260-00-001
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [ValidateSet('table','json','csv')] [string]$Format = 'table',
  [int]$MaxDepth = 20,
  [string]$Email = $env:ARENA_EMAIL,
  [string]$WorkspaceId = $env:ARENA_WORKSPACE_ID,
  [string]$BaseUrl = 'https://api.arenasolutions.com/v1'
)

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

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

$password = $env:ARENA_PASSWORD
if ([string]::IsNullOrWhiteSpace($Email))       { throw 'Set ARENA_EMAIL or pass -Email.' }
if ([string]::IsNullOrWhiteSpace($password))    { throw 'Set ARENA_PASSWORD.' }
if ([string]::IsNullOrWhiteSpace($WorkspaceId)) { throw 'Set ARENA_WORKSPACE_ID or pass -WorkspaceId.' }

$apiRoot = $BaseUrl.TrimEnd('/')

$loginBody = @{ email = $Email; password = $password; workspaceId = [long]$WorkspaceId } | ConvertTo-Json
try { $login = Invoke-ArenaJson -Uri "$apiRoot/login" -Method Post -Body $loginBody }
catch { throw "Arena login failed: $($_.Exception.Message)" }

$sid = $login.arenaSessionId
if ([string]::IsNullOrWhiteSpace($sid)) { $sid = $login.arena_session_id }
if ([string]::IsNullOrWhiteSpace($sid)) { throw 'Login response did not include arenaSessionId.' }

$headers = @{ arena_session_id = $sid; 'Content-Type' = 'application/json' }

# ---- resolve top item ----

$pnEnc = [Uri]::EscapeDataString($PartNumber.Trim())
$lookup = Invoke-ArenaJson -Uri "$apiRoot/items?number=$pnEnc&limit=1" -Headers $headers
if (-not $lookup.results -or $lookup.results.Count -eq 0) {
  throw "No item found for number '$($PartNumber.Trim())'."
}
$top = $lookup.results[0]

# ---- walk ----

$pcbaPattern  = '^(?<prefix>\d{3})-(?<asm>10|15)-(?<suffix>\d+)$'
$visited      = New-Object 'System.Collections.Generic.HashSet[string]'
$pcbas        = @{}   # part number -> record
$apiCallCount = 0

function Walk-Bom {
  param(
    [string]$Guid,
    [string[]]$PathChain,   # array of part numbers from root down to this item's parent
    [int]$Depth
  )
  if ($Depth -gt $MaxDepth)         { return }
  if ([string]::IsNullOrWhiteSpace($Guid)) { return }
  if (-not $script:visited.Add($Guid))     { return }

  try {
    $script:apiCallCount++
    $bom = Invoke-ArenaJson -Uri "$apiRoot/items/$Guid/bom" -Headers $headers
  } catch {
    Write-Warning "BOM fetch failed at depth $Depth, guid=$Guid : $($_.Exception.Message)"
    return
  }
  if (-not $bom -or -not $bom.results) { return }

  foreach ($line in $bom.results) {
    $child = $line.item
    if (-not $child) { continue }
    $childPn = $child.number

    $m = [regex]::Match($childPn, $pcbaPattern)
    if ($m.Success -and $m.Groups['prefix'].Value -ne '098') {
      if (-not $script:pcbas.ContainsKey($childPn)) {
        $script:pcbas[$childPn] = [pscustomobject]@{
          PartNumber  = $childPn
          Name        = $child.name
          Revision    = $child.revisionNumber
          Guid        = $child.guid
          Occurrences = 0
          FirstParentPath = ($PathChain -join ' > ')
          AppUrl      = ($child.url.app)
        }
      }
      $script:pcbas[$childPn].Occurrences++
    }

    Walk-Bom -Guid $child.guid -PathChain ($PathChain + @($childPn)) -Depth ($Depth + 1)
  }
}

Walk-Bom -Guid $top.guid -PathChain @($top.number) -Depth 0

$ordered = $pcbas.Values | Sort-Object PartNumber

switch ($Format) {
  'json' {
    [pscustomobject]@{
      Top = [pscustomobject]@{
        Number   = $top.number
        Name     = $top.name
        Revision = $top.revisionNumber
        Guid     = $top.guid
      }
      ApiCalls   = $apiCallCount
      Found      = $ordered.Count
      Assemblies = $ordered
    } | ConvertTo-Json -Depth 6
  }
  'csv' {
    $ordered | ConvertTo-Csv -NoTypeInformation
  }
  default {
    Write-Output ("Top: {0}  {1}  Rev {2}" -f $top.number, $top.name, $top.revisionNumber)
    Write-Output ("Visited {0} item(s) via {1} BOM call(s); found {2} unique PCB assembly part(s)." -f `
      $script:visited.Count, $apiCallCount, $ordered.Count)
    Write-Output ''
    if ($ordered.Count -eq 0) {
      Write-Output "No PCBAs (PPP-10-YYY or PPP-15-YYY, prefix != 098) found in this BOM."
    } else {
      $ordered |
        Select-Object PartNumber, Revision, Occurrences, Name, FirstParentPath |
        Format-Table -AutoSize |
        Out-String -Width 240
    }
  }
}

if ($pcbas.Count -eq 0) { exit 2 } else { exit 0 }
