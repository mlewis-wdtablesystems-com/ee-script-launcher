<#
.SYNOPSIS
  Minimal window: look up an Arena item and its BOM by part number.

.DESCRIPTION
  Reads ARENA_EMAIL, ARENA_PASSWORD, and ARENA_WORKSPACE_ID from the environment,
  then uses the Arena REST API (login, GET /items, GET /items/{guid}/bom).

.EXAMPLE
  $env:ARENA_EMAIL = 'you@company.com'
  $env:ARENA_PASSWORD = '...'
  $env:ARENA_WORKSPACE_ID = '896309030'
  .\Show-ArenaPlmDashboard.ps1

.EXAMPLE
  $env:ARENA_PASSWORD = '...'
  .\Show-ArenaPlmDashboard.ps1 -Email 'you@company.com' -WorkspaceId '896309030'
#>
param(
  [string]$Email = $env:ARENA_EMAIL,
  [string]$WorkspaceId = $env:ARENA_WORKSPACE_ID,
  [string]$BaseUrl = 'https://api.arenasolutions.com/v1'
)

$ErrorActionPreference = 'Stop'

$apiRoot = $BaseUrl.TrimEnd('/')

Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
Add-Type -AssemblyName System.Drawing -ErrorAction Stop

if ([string]::IsNullOrWhiteSpace($Email) -or
    [string]::IsNullOrWhiteSpace($env:ARENA_PASSWORD) -or
    [string]::IsNullOrWhiteSpace($WorkspaceId)) {
  [System.Windows.Forms.MessageBox]::Show(
    "Provide email and workspace (parameters or ARENA_EMAIL / ARENA_WORKSPACE_ID) and set ARENA_PASSWORD.",
    'Arena PLM'
  ) | Out-Null
  exit 1
}

function Get-ArenaSessionHeaders {
  $password = $env:ARENA_PASSWORD
  if ([string]::IsNullOrWhiteSpace($password)) {
    throw 'Set ARENA_PASSWORD.'
  }
  $loginBody = @{
    email       = $Email
    password    = $password
    workspaceId = [long]$WorkspaceId
  } | ConvertTo-Json
  $login = Invoke-RestMethod -Uri "$apiRoot/login" -Method Post `
    -ContentType 'application/json' -Body $loginBody
  $sid = $login.arenaSessionId
  if ([string]::IsNullOrWhiteSpace($sid)) { $sid = $login.arena_session_id }
  if ([string]::IsNullOrWhiteSpace($sid)) { throw 'Login did not return a session id.' }
  return @{
    arena_session_id = $sid
    'Content-Type'   = 'application/json'
  }
}

function Get-ArenaItemRecord {
  param(
    [Parameter(Mandatory = $true)] [hashtable]$Headers,
    [Parameter(Mandatory = $true)] [string]$PartNumber
  )
  $pn = $PartNumber.Trim()
  $pnEnc = [Uri]::EscapeDataString($pn)
  $uri = "$apiRoot/items?number=$pnEnc&limit=50&responseview=full"
  $resp = Invoke-RestMethod -Uri $uri -Method Get -Headers $Headers
  if (-not $resp.results -or $resp.results.Count -eq 0) { return $null }
  return $resp.results[0]
}

function Format-ItemSummary {
  param([Parameter(Mandatory = $true)] $Item)
  $lp = $Item.lifecyclePhase
  $cat = $Item.category
  $phaseName = if ($lp.name) { $lp.name } else { '' }
  $catName = if ($cat.name) { $cat.name } else { '' }
  $lines = @(
    "Number:           $($Item.number)",
    "Name:             $($Item.name)",
    "Revision:         $($Item.revisionNumber)  ($($Item.revisionStatus))",
    "Lifecycle:        $phaseName",
    "Category:         $catName",
    "Assembly type:    $($Item.assemblyType)",
    "UOM:              $($Item.uom)",
    "Effective:        $($Item.effectiveDateTime)",
    "GUID:             $($Item.guid)",
    "API:              $($Item.url.api)"
  )
  return ($lines -join [Environment]::NewLine)
}

function Format-BomTable {
  param([Parameter(Mandatory = $true)] $BomResponse)
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine("Qty`tPart number`tRev`tName`tRefDes")
  foreach ($row in $BomResponse.results) {
    $it = $row.item
    $rev = $it.revisionNumber
    $ref = if ($null -ne $row.refDes) { $row.refDes } else { '' }
    [void]$sb.AppendLine("$($row.quantity)`t$($it.number)`t$rev`t$($it.name)`t$ref")
  }
  [void]$sb.AppendLine()
  [void]$sb.AppendLine("$($BomResponse.count) BOM line(s).")
  return $sb.ToString()
}

function Invoke-GetItemClick {
  param($TextBoxPart, $TextBoxItem, $TextBoxBom)
  $pn = $TextBoxPart.Text
  if ([string]::IsNullOrWhiteSpace($pn)) {
    [System.Windows.Forms.MessageBox]::Show('Enter a part number.') | Out-Null
    return
  }
  try {
    $h = Get-ArenaSessionHeaders
    $rec = Get-ArenaItemRecord -Headers $h -PartNumber $pn
    if (-not $rec) {
      $TextBoxItem.Text = "No item found for number: $($pn.Trim())"
      return
    }
    $TextBoxItem.Text = Format-ItemSummary -Item $rec
  } catch {
    $TextBoxItem.Text = "Error: $($_.Exception.Message)"
  }
}

function Invoke-GetBomClick {
  param($TextBoxPart, $TextBoxItem, $TextBoxBom)
  $pn = $TextBoxPart.Text
  if ([string]::IsNullOrWhiteSpace($pn)) {
    [System.Windows.Forms.MessageBox]::Show('Enter a part number.') | Out-Null
    return
  }
  try {
    $h = Get-ArenaSessionHeaders
    $rec = Get-ArenaItemRecord -Headers $h -PartNumber $pn
    if (-not $rec) {
      $TextBoxBom.Text = "No item found for number: $($pn.Trim())"
      return
    }
    $guid = $rec.guid
    $bom = Invoke-RestMethod -Uri "$apiRoot/items/$guid/bom" -Method Get -Headers $h
    $TextBoxBom.Text = Format-BomTable -BomResponse $bom
  } catch {
    $TextBoxBom.Text = "Error: $($_.Exception.Message)"
  }
}

[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Arena PLM'
$form.Size = New-Object System.Drawing.Size(720, 520)
$form.MinimumSize = New-Object System.Drawing.Size(480, 360)
$form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

$lblPart = New-Object System.Windows.Forms.Label
$lblPart.Text = 'Part number'
$lblPart.Location = New-Object System.Drawing.Point(12, 14)
$lblPart.AutoSize = $true

$txtPart = New-Object System.Windows.Forms.TextBox
$txtPart.Location = New-Object System.Drawing.Point(100, 11)
$txtPart.Width = 320
$txtPart.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

$btnItem = New-Object System.Windows.Forms.Button
$btnItem.Text = 'Item info'
$btnItem.Width = 88
$btnItem.Height = 24
$btnItem.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$btnItem.Add_Click({ Invoke-GetItemClick $txtPart $txtItem $txtBom })

$btnBom = New-Object System.Windows.Forms.Button
$btnBom.Text = 'Get BOM'
$btnBom.Width = 88
$btnBom.Height = 24
$btnBom.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$btnBom.Add_Click({ Invoke-GetBomClick $txtPart $txtItem $txtBom })

$lblItem = New-Object System.Windows.Forms.Label
$lblItem.Text = 'Item'
$lblItem.AutoSize = $true

$txtItem = New-Object System.Windows.Forms.TextBox
$txtItem.Height = 140
$txtItem.Multiline = $true
$txtItem.ReadOnly = $true
$txtItem.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtItem.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtItem.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$lblBom = New-Object System.Windows.Forms.Label
$lblBom.Text = 'BOM'
$lblBom.AutoSize = $true

$txtBom = New-Object System.Windows.Forms.TextBox
$txtBom.Multiline = $true
$txtBom.ReadOnly = $true
$txtBom.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$txtBom.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtBom.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

$form.Controls.AddRange(@($lblPart, $txtPart, $btnItem, $btnBom, $lblItem, $txtItem, $lblBom, $txtBom))

$layoutDashboard = {
  $m = 12
  $pw = $form.ClientSize.Width
  $ph = $form.ClientSize.Height
  $btnBom.Top = 9
  $btnBom.Left = $pw - $m - $btnBom.Width
  $btnItem.Top = 9
  $btnItem.Left = $btnBom.Left - $btnItem.Width - 8
  $txtPart.Width = [Math]::Max(120, $btnItem.Left - $txtPart.Left - 8)

  $lblItem.Left = $m
  $lblItem.Top = 44
  $txtItem.Left = $m
  $txtItem.Top = $lblItem.Bottom + 6
  $txtItem.Width = $pw - 2 * $m
  $remaining = $ph - $txtItem.Top - $m
  $itemH = [int]($remaining * 0.38)
  if ($itemH -lt 88) { $itemH = 88 }
  if ($itemH -gt 240) { $itemH = 240 }
  $txtItem.Height = $itemH

  $lblBom.Left = $m
  $lblBom.Top = $txtItem.Bottom + 10
  $txtBom.Left = $m
  $txtBom.Top = $lblBom.Bottom + 4
  $txtBom.Width = $pw - 2 * $m
  $txtBom.Height = [Math]::Max(96, $ph - $txtBom.Top - $m)
}

$form.Add_Resize($layoutDashboard)
$form.Add_Shown({ & $layoutDashboard })

[void][System.Windows.Forms.Application]::Run($form)
