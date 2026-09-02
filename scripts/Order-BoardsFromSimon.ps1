<#
.SYNOPSIS
  Draft an Outlook email to Simon Graham to order bare boards, attaching Fabrication and Gerber zips.

.DESCRIPTION
  1. Resolves the Altium production folder (Open-AltiumPartFolder.ps1 -Open ResolveJson).
  2. Locates or builds:
       Fabrication <PPP-20/25-YYY> Rev <letter>.zip
       Gerbers <PPP-20/25-YYY> Rev <letter>.zip  (revision taken from the Fabrication zip name)
  3. Attaches the single .pdf from the production folder's PFS subfolder.
  4. Attaches the schematic PDF from the main design folder: either the PDF next to the
     newest *.PrjPcb (same base name, excluding Archive), or a PDF under a "PDF" subfolder
     titled "Schecmatics of ..." / "Schematics of ...".
  5. Opens a draft Outlook message to sgraham@wdtablesystems.com
     (CC: bthakkar@wdtablesystems.com, rohit.maurya@wdtablesystems.com,
      gahamad@wdtablesystems.com, sks@wdtablesystems.com).

  Uses Classic Outlook COM when available. When New Outlook is running (no COM), builds an
  .eml draft and opens it with olk.exe so the message appears in your current Outlook app.

.PARAMETER PartNumber
  PPP-XX-XXX with optional folder suffix (same rules as Open-AltiumPartFolder.ps1).

.PARAMETER Quantity
  Board quantity for the order email (number or text, e.g. "Marcos to decide").

.PARAMETER Description
  Optional extra sentence in the order email (shown after the order line).

.PARAMETER Urgency
  high - "Please expedite as fast as possible."
  low  - "These are not super urgent"
  unstated - omit the urgency line from the email
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [Parameter(Mandatory = $true)] [string]$Quantity,
  [Parameter(Mandatory = $false)] [string]$Description = '',
  [Parameter(Mandatory = $true)]
  [ValidateSet('high', 'low', 'unstated')]
  [string]$Urgency
)

# Outlook COM requires an STA thread; the launcher invokes PowerShell without -STA.
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
  $staArgs = @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', $MyInvocation.MyCommand.Path,
    '-PartNumber', $PartNumber,
    '-Quantity', $Quantity,
    '-Description', $Description,
    '-Urgency', $Urgency
  )
  & powershell.exe @staArgs
  exit $LASTEXITCODE
}

$ErrorActionPreference = 'Stop'
$OutputEncoding = [System.Text.Encoding]::UTF8
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

$scriptsDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$openScript = Join-Path $scriptsDir 'Open-AltiumPartFolder.ps1'
$explorerHelpers = Join-Path $scriptsDir 'ExplorerWindowHelpers.ps1'
if (-not (Test-Path -LiteralPath $openScript)) { throw "Required script missing: $openScript" }
if (Test-Path -LiteralPath $explorerHelpers) { . $explorerHelpers }

$pn = $PartNumber.Trim()
$qty = $Quantity.Trim()
$descriptionText = $Description.Trim()
if ([string]::IsNullOrWhiteSpace($qty)) {
  throw 'Quantity is required.'
}

$urgencyLine = switch ($Urgency) {
  'high' { 'Please expedite as fast as possible.' }
  'low'  { 'These are not super urgent' }
  'unstated' { $null }
}

function ConvertTo-OrderBoardHtmlBody {
  param([string[]]$Lines)
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append('<html><head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head><body>')
  [void]$sb.Append('<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;">')
  foreach ($line in $Lines) {
    if ($line -eq '') {
      [void]$sb.Append('<div><br></div>')
    }
    else {
      $escaped = [System.Net.WebUtility]::HtmlEncode($line)
      [void]$sb.Append("<div>$escaped</div>")
    }
  }
  [void]$sb.Append('</div></body></html>')
  return $sb.ToString()
}

function New-OrderBoardBodyLines {
  param(
    [string]$PcbPartNumber,
    [string]$Revision,
    [string]$Description,
    [string]$Quantity,
    [AllowNull()]
    [string]$UrgencyText
  )
  $partRev = if ($Revision) { "$PcbPartNumber Rev $Revision" } else { $PcbPartNumber }
  $lines = New-Object System.Collections.Generic.List[string]
  [void]$lines.Add('Hi Simon,')
  [void]$lines.Add('')
  [void]$lines.Add("Can you please order $partRev.")
  if ($Description) {
    [void]$lines.Add($Description)
  }
  [void]$lines.Add("Quantity = $Quantity")
  if ($UrgencyText) { [void]$lines.Add($UrgencyText) }
  [void]$lines.Add('')
  [void]$lines.Add('Thanks,')
  [void]$lines.Add('')
  [void]$lines.Add('Marshall')
  [void]$lines.Add('Generated with Script')
  return @($lines)
}

function Get-ResolveJsonFromOpenScript {
  param([string]$PartLookup)
  $resolveOut = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
    -File $openScript -PartNumber $PartLookup -Open ResolveJson 2>&1
  if ($LASTEXITCODE -ne 0) {
    throw "Open-AltiumPartFolder.ps1 -Open ResolveJson failed (exit $LASTEXITCODE):`n$($resolveOut -join [Environment]::NewLine)"
  }
  $joined = ($resolveOut | ForEach-Object { "$_" }) -join "`n"
  $rxJson = [regex]'---BEGIN-RESOLVE-JSON---\s*(?<j>\{.*\})\s*---END-RESOLVE-JSON---'
  $jm = $rxJson.Match($joined)
  if (-not $jm.Success) {
    throw "Could not find resolve JSON markers. Output:`n$joined"
  }
  return ($jm.Groups['j'].Value | ConvertFrom-Json)
}

function Get-AltiumZipBuildFailureMessage {
  param(
    [string]$OutputText,
    [string]$Label,
    [int]$ExitCode
  )
  $lines = @(
    $OutputText -split "`r?`n" |
      ForEach-Object { "$_".TrimEnd() } |
      Where-Object { $_.Trim().Length -gt 0 }
  )

  $errorLines = @($lines | Where-Object { $_ -match '\bERROR:' })
  if ($errorLines.Count -gt 0) {
    return ("Build $Label failed (exit $ExitCode):`n" + ($errorLines -join "`n"))
  }

  $aborted = @($lines | Where-Object { $_ -match 'zip aborted:' })
  if ($aborted.Count -gt 0) {
    return ($aborted[-1])
  }

  $throwLines = @($lines | Where-Object {
      $_ -match '^(Fabrication zip aborted|Gerber zip aborted|None of the required folders|Production document|PFS folder|Could not determine revision|No matching folder found|Missing folder:)'
    })
  if ($throwLines.Count -gt 0) {
    return ("Build $Label failed (exit $ExitCode):`n" + ($throwLines -join "`n"))
  }

  # PowerShell nested invocation: "Open-AltiumPartFolder.ps1 : <message>"
  $psError = @($lines | Where-Object { $_ -match '^Open-AltiumPartFolder\.ps1\s*:\s*(.+)$' })
  if ($psError.Count -gt 0) {
    $msg = ($psError[-1] -replace '^Open-AltiumPartFolder\.ps1\s*:\s*', '').Trim()
    if ($msg) {
      return ("Build $Label failed (exit $ExitCode): $msg")
    }
  }

  if ($lines.Count -gt 0) {
    return ("Build $Label failed (exit $ExitCode):`n" + ($lines -join "`n"))
  }

  return "Build $Label failed (exit $ExitCode)."
}

function Ensure-AltiumZip {
  param(
    [string]$PartLookup,
    [ValidateSet('FabricationZip', 'GerberZip')]
    [string]$OpenMode,
    [string]$ZipPath,
    [string]$Label
  )
  if (Test-Path -LiteralPath $ZipPath) {
    Write-Host "  Using existing $Label`: $ZipPath"
    return
  }
  Write-Host ''
  Write-Host "=== $Label not found; building via Open-AltiumPartFolder -Open $OpenMode ==="
  $openArgs = @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass',
    '-File', $openScript,
    '-PartNumber', $PartLookup,
    '-Open', $OpenMode
  )
  $out = & powershell.exe @openArgs 2>&1
  $text = ($out | ForEach-Object { "$_" }) -join "`n"
  if ($text.Trim()) { Write-Host $text }
  if ($LASTEXITCODE -ne 0) {
    throw (Get-AltiumZipBuildFailureMessage -OutputText $text -Label $Label -ExitCode $LASTEXITCODE)
  }
  if (-not (Test-Path -LiteralPath $ZipPath)) {
    throw (Get-AltiumZipBuildFailureMessage -OutputText $text -Label $Label -ExitCode 1) `
      + "`n$Label still missing after build: $ZipPath"
  }
  Write-Host "  Created: $ZipPath"
}

function Find-FabricationZipInProdFolder {
  param(
    [string]$ProdDocDir,
    [string]$PcbPartNumber
  )
  if (-not (Test-Path -LiteralPath $ProdDocDir -PathType Container)) { return $null }
  $candidates = @(Get-ChildItem -LiteralPath $ProdDocDir -File -ErrorAction SilentlyContinue |
    Where-Object {
      $_.Extension -ieq '.zip' -and
      $_.Name -imatch '^Fabrication\s+.+\s+Rev\s+[A-Za-z0-9]+\.zip$'
    })
  if ($PcbPartNumber) {
    $forPart = @($candidates | Where-Object { $_.Name -imatch [regex]::Escape($PcbPartNumber) })
    if ($forPart.Count -gt 0) { $candidates = $forPart }
  }
  if ($candidates.Count -eq 0) { return $null }
  return ($candidates | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1)
}

function Get-RevisionFromFabricationZipFileName {
  param([string]$ZipFileName)
  if ($ZipFileName -match '(?i)\bRev\s+([A-Za-z0-9]+)\.zip$') {
    return $matches[1]
  }
  return $null
}

function Ensure-FabricationZipForOrderBoard {
  param(
    [string]$PartLookup,
    [string]$ProdDocDir,
    [string]$PcbPartNumber
  )
  $found = Find-FabricationZipInProdFolder -ProdDocDir $ProdDocDir -PcbPartNumber $PcbPartNumber
  if ($found) {
    Write-Host "  Using existing Fabrication zip: $($found.FullName)"
    return $found.FullName
  }

  Write-Host ''
  Write-Host '=== Fabrication zip not found; building via Open-AltiumPartFolder -Open FabricationZip ==='
  $openArgs = @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass',
    '-File', $openScript,
    '-PartNumber', $PartLookup,
    '-Open', 'FabricationZip'
  )
  $out = & powershell.exe @openArgs 2>&1
  $text = ($out | ForEach-Object { "$_" }) -join "`n"
  if ($text.Trim()) { Write-Host $text }
  if ($LASTEXITCODE -ne 0) {
    throw (Get-AltiumZipBuildFailureMessage -OutputText $text -Label 'Fabrication zip' -ExitCode $LASTEXITCODE)
  }

  $found = Find-FabricationZipInProdFolder -ProdDocDir $ProdDocDir -PcbPartNumber $PcbPartNumber
  if (-not $found) {
    throw "Fabrication zip still missing after build in `"$ProdDocDir`"."
  }
  Write-Host "  Created: $($found.FullName)"
  return $found.FullName
}

function Get-SchematicPdfBesidePrjPcb {
  param([string]$SearchRoot)
  if ([string]::IsNullOrWhiteSpace($SearchRoot) -or -not (Test-Path -LiteralPath $SearchRoot -PathType Container)) {
    throw "Design folder not found for schematic PDF: $SearchRoot"
  }

  $allPrj = @(Get-ChildItem -LiteralPath $SearchRoot -Recurse -File -Filter '*.PrjPcb' -ErrorAction SilentlyContinue |
    Where-Object {
      -not (($_.FullName -split '[\\/]') -contains 'Archive')
    })
  if ($allPrj.Count -eq 0) {
    throw "No .PrjPcb file found under `"$SearchRoot`" (paths under an `"Archive`" folder are ignored)."
  }
  $prj = $allPrj | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
  $prjDir = $prj.DirectoryName

  # 1) Same base name as .PrjPcb, beside the project file
  $beside = @(Get-ChildItem -LiteralPath $prjDir -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Extension -ieq '.pdf' -and $_.BaseName -ieq $prj.BaseName })
  if ($beside.Count -gt 0) {
    return ($beside | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1).FullName
  }

  # 2) PDF subfolder: "Schecmatics of ..." (common typo) or "Schematics of ..."
  $pdfDirs = @(
    (Join-Path $prjDir 'PDF')
    (Join-Path $SearchRoot 'PDF')
  ) | Select-Object -Unique
  $titleHits = @()
  foreach ($pdfDir in $pdfDirs) {
    if (-not (Test-Path -LiteralPath $pdfDir -PathType Container)) { continue }
    $titleHits += @(Get-ChildItem -LiteralPath $pdfDir -File -ErrorAction SilentlyContinue |
      Where-Object {
        $_.Extension -ieq '.pdf' -and
        ($_.BaseName -match '(?i)^Schec?matics\s+of\b')
      })
  }
  if ($titleHits.Count -gt 0) {
    return ($titleHits | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1).FullName
  }

  $looked = @(
    "`"$($prj.BaseName).pdf`" beside `"$prjDir`""
    '"Schecmatics of …" / "Schematics of …" under a PDF subfolder'
  ) -join '; '
  throw "Schematic PDF not found. Looked for: $looked."
}

function Get-PfsPdfPath {
  param([string]$ProdDocDir)
  $pfsDir = Join-Path $ProdDocDir 'PFS'
  if (-not (Test-Path -LiteralPath $pfsDir -PathType Container)) {
    throw "PFS folder not found: $pfsDir"
  }
  $pdfs = @(Get-ChildItem -LiteralPath $pfsDir -File -Filter '*.pdf' -ErrorAction SilentlyContinue)
  if ($pdfs.Count -eq 0) {
    throw "PFS folder must contain a .pdf file (none found in $pfsDir)."
  }
  if ($pdfs.Count -gt 1) {
    $listed = ($pdfs | Sort-Object Name | ForEach-Object { $_.Name }) -join ', '
    throw "PFS folder must contain exactly one .pdf file (found $($pdfs.Count): $listed)."
  }
  return $pdfs[0].FullName
}

function Get-NewOutlookExePath {
  $candidates = @(
    (Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\olk.exe'),
    (Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\Microsoft.OutlookForWindows_8wekyb3d8bbwe\olk.exe')
  )
  foreach ($p in $candidates) {
    if (Test-Path -LiteralPath $p) { return $p }
  }
  return $null
}

function Get-ClassicOutlookExePath {
  $candidates = @(
    (Join-Path ${env:ProgramFiles} 'Microsoft Office\Root\Office16\OUTLOOK.EXE'),
    (Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\Root\Office16\OUTLOOK.EXE')
  )
  foreach ($p in $candidates) {
    if (Test-Path -LiteralPath $p) { return $p }
  }
  return $null
}

function Invoke-FocusOutlookWindow {
  try {
    if (Get-Command Show-AppWindowToFront -ErrorAction SilentlyContinue) {
      Show-AppWindowToFront -TitleHint 'Outlook'
      return
    }
    $wsh = New-Object -ComObject 'WScript.Shell'
    [void]$wsh.AppActivate('Outlook')
  }
  catch {}
}

function Get-OutlookComApplication {
  Invoke-FocusOutlookWindow
  foreach ($attempt in 1..8) {
    try {
      $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
      Write-Host "Connected to running Classic Outlook (attempt $attempt)."
      return $app
    }
    catch {
      if ($attempt -eq 1) {
        Write-Host 'Trying to connect to Classic Outlook COM...'
      }
      Start-Sleep -Milliseconds 750
    }
  }
  return $null
}

function Get-MimeTypeForPath {
  param([string]$Path)
  switch ([IO.Path]::GetExtension($Path).ToLowerInvariant()) {
    '.pdf' { return 'application/pdf' }
    '.zip' { return 'application/zip' }
    default { return 'application/octet-stream' }
  }
}

function New-OrderBoardEmlFile {
  param(
    [string]$To,
    [string]$Cc = '',
    [string]$Subject,
    [string]$HtmlBody,
    [string[]]$AttachmentPaths,
    [string]$OutputPath
  )

  $boundary = "----=_Part_$([Guid]::NewGuid().ToString('N'))"
  $lines = New-Object System.Collections.Generic.List[string]
  [void]$lines.Add('X-Unsent: 1')
  [void]$lines.Add("To: $To")
  if ($Cc) { [void]$lines.Add("Cc: $Cc") }
  [void]$lines.Add("Subject: $Subject")
  [void]$lines.Add('MIME-Version: 1.0')
  [void]$lines.Add("Content-Type: multipart/mixed; boundary=`"$boundary`"")
  [void]$lines.Add('')
  [void]$lines.Add("--$boundary")
  [void]$lines.Add('Content-Type: text/html; charset="utf-8"')
  [void]$lines.Add('Content-Transfer-Encoding: 8bit')
  [void]$lines.Add('')
  foreach ($line in ($HtmlBody -split "`r?`n")) { [void]$lines.Add($line) }

  foreach ($att in $AttachmentPaths) {
    if (-not (Test-Path -LiteralPath $att)) {
      throw "Attachment not found for EML: $att"
    }
    $fileName = [IO.Path]::GetFileName($att)
    $mimeType = Get-MimeTypeForPath -Path $att
    $b64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($att))
    [void]$lines.Add("--$boundary")
    [void]$lines.Add("Content-Type: $mimeType; name=`"$fileName`"")
    [void]$lines.Add('Content-Transfer-Encoding: base64')
    [void]$lines.Add("Content-Disposition: attachment; filename=`"$fileName`"")
    [void]$lines.Add('')
    for ($i = 0; $i -lt $b64.Length; $i += 76) {
      $len = [Math]::Min(76, $b64.Length - $i)
      [void]$lines.Add($b64.Substring($i, $len))
    }
  }
  [void]$lines.Add("--$boundary--")
  [void]$lines.Add('')

  # LF-only body; system ANSI encoding helps New Outlook show attachments reliably.
  $content = ($lines -join "`n")
  [IO.File]::WriteAllText($OutputPath, $content, [Text.Encoding]::Default)
}

function Open-OutlookEmlDraft {
  param([string]$EmlPath)

  $olk = Get-NewOutlookExePath
  if ($olk) {
    Write-Host "Opening draft in New Outlook: $olk"
    Start-Process -FilePath $olk -ArgumentList $EmlPath | Out-Null
    return
  }

  $classic = Get-ClassicOutlookExePath
  if ($classic) {
    Write-Host "Opening draft in Classic Outlook (/eml): $classic"
    Start-Process -FilePath $classic -ArgumentList '/eml', $EmlPath | Out-Null
    return
  }

  Write-Host "Opening draft via default .eml handler: $EmlPath"
  Start-Process -FilePath $EmlPath | Out-Null
}

function Set-OutlookMailHtmlBodyWithoutSignature {
  param(
    [Parameter(Mandatory = $true)] $MailItem,
    [Parameter(Mandatory = $true)] [string]$HtmlBody
  )
  # olFormatHTML = 2 — rich compose (images, links, @ mentions); replace HTMLBody after
  # inspector init so the default signature is not kept.
  $MailItem.BodyFormat = 2
  $null = $MailItem.GetInspector()
  Start-Sleep -Milliseconds 400
  $MailItem.HTMLBody = $HtmlBody
}

function Open-OrderBoardOutlookDraft {
  param(
    [string]$To,
    [string]$Cc = '',
    [string]$Subject,
    [string]$HtmlBody,
    [string[]]$AttachmentPaths
  )

  $outlookRunning = [bool](Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
  $outlook = Get-OutlookComApplication
  if ($outlook) {
    Write-Host '=== Opening Outlook draft (Classic COM, HTML) ==='
    $mail = $null
    try {
      $mail = $outlook.CreateItem(0)
      $mail.To = $To
      if ($Cc) { $mail.CC = $Cc }
      $mail.Subject = $Subject
      foreach ($att in $AttachmentPaths) {
        [void]$mail.Attachments.Add($att)
      }
      Set-OutlookMailHtmlBodyWithoutSignature -MailItem $mail -HtmlBody $HtmlBody
      $mail.Display()
      Write-Host 'Draft opened in Classic Outlook (review and send when ready).'
      return
    }
    finally {
      if ($mail) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) }
      if ($outlook) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) }
    }
  }

  Write-Host '=== Opening Outlook draft (.eml, HTML) ==='
  if ($outlookRunning -and (Get-NewOutlookExePath)) {
    Write-Host 'New Outlook is running (no COM). Using olk.exe with an .eml draft instead of starting Classic Outlook.'
  }
  elseif ($outlookRunning) {
    Write-Host 'Outlook is running but COM is unavailable. Using an .eml draft instead of starting a second Outlook.'
  }

  $emlPath = Join-Path ([IO.Path]::GetTempPath()) ("wdts-order-board-$([Guid]::NewGuid().ToString('N')).eml")
  New-OrderBoardEmlFile -To $To -Cc $Cc -Subject $Subject -HtmlBody $HtmlBody -AttachmentPaths $AttachmentPaths -OutputPath $emlPath
  Write-Host "Draft file: $emlPath"
  Open-OutlookEmlDraft -EmlPath $emlPath
  Write-Host 'Draft opened in Outlook (review and send when ready).'
}

Write-Host '=== Altium resolve ==='
$resolved = Get-ResolveJsonFromOpenScript -PartLookup $pn
$prodDocDir = $resolved.prodDocDir
if ([string]::IsNullOrWhiteSpace($prodDocDir)) {
  throw 'ResolveJson did not return prodDocDir (no production folder under any revision?).'
}
$pcbPart = $resolved.pcbPartNumber
if ([string]::IsNullOrWhiteSpace($pcbPart)) {
  throw 'ResolveJson did not return pcbPartNumber.'
}
$designDir = $resolved.targetDir
if ([string]::IsNullOrWhiteSpace($designDir)) {
  $designDir = $resolved.partDir
}
if ([string]::IsNullOrWhiteSpace($designDir)) {
  throw 'ResolveJson did not return targetDir or partDir (needed for schematic PDF).'
}

Write-Host ''
Write-Host '=== Attachments ==='
$fabricationZip = Ensure-FabricationZipForOrderBoard `
  -PartLookup $pn `
  -ProdDocDir $prodDocDir `
  -PcbPartNumber $pcbPart

$revFromFab = Get-RevisionFromFabricationZipFileName -ZipFileName (Split-Path -Leaf $fabricationZip)
if ([string]::IsNullOrWhiteSpace($revFromFab)) {
  throw "Could not parse revision from fabrication zip name: $fabricationZip"
}
$revLetterOnly = ($revFromFab -replace '\d+$', '')
if ([string]::IsNullOrWhiteSpace($revLetterOnly)) { $revLetterOnly = $revFromFab }

$gerberZip = Join-Path $prodDocDir ("Gerbers $pcbPart Rev $revLetterOnly.zip")
$pfsPdf = Get-PfsPdfPath -ProdDocDir $prodDocDir
$schematicPdf = Get-SchematicPdfBesidePrjPcb -SearchRoot $designDir

Write-Host "  Production document folder: $prodDocDir"
Write-Host "  Design folder (schematic):  $designDir"
Write-Host "  Lookup part (Altium):       $pn"
Write-Host "  Bare-board part (email):    $pcbPart"
Write-Host "  Revision (from Fabrication zip): Rev $revLetterOnly"
Write-Host "  Fabrication zip:            $fabricationZip"
Write-Host "  Gerber zip:                 $gerberZip"
Write-Host "  PFS pdf:                    $pfsPdf"
Write-Host "  Schematic pdf:              $schematicPdf"

Ensure-AltiumZip -PartLookup $pn -OpenMode GerberZip -ZipPath $gerberZip -Label 'Gerber zip'
Write-Host "  Using PFS pdf: $pfsPdf"
Write-Host "  Using schematic pdf: $schematicPdf"

$bodyLines = New-OrderBoardBodyLines `
  -PcbPartNumber $pcbPart `
  -Revision $revLetterOnly `
  -Description $descriptionText `
  -Quantity $qty `
  -UrgencyText $urgencyLine
$htmlBody = ConvertTo-OrderBoardHtmlBody -Lines $bodyLines

Open-OrderBoardOutlookDraft `
  -To 'sgraham@wdtablesystems.com' `
  -Cc 'bthakkar@wdtablesystems.com; rohit.maurya@wdtablesystems.com; gahamad@wdtablesystems.com; sks@wdtablesystems.com' `
  -Subject "Order $pcbPart Rev $revLetterOnly" `
  -HtmlBody $htmlBody `
  -AttachmentPaths @($fabricationZip, $gerberZip, $pfsPdf, $schematicPdf)

exit 0
