<#
.SYNOPSIS
  Draft an Outlook email to Neel with _DRAFT fabrication and Pick & Place zips.

.DESCRIPTION
  Resolves the Altium production folder, attaches the newest Fabrication *_DRAFT.zip
  and Pick Place *_DRAFT.zip, and attaches a schematic PDF when one is found.
  If neither draft zip exists, exits 14 so the launcher can ask whether to send anyway.
  A missing schematic is a warning only; the email is still drafted.

.PARAMETER PartNumber
  PPP-XX-XXX with optional folder suffix (same rules as Open-AltiumPartFolder.ps1).

.PARAMETER AllowMissingDraftFiles
  Draft the email even when no _DRAFT fabrication or Pick & Place zip exists.
  The launcher sets this after "Send anyway?".
#>
param(
  [Parameter(Mandatory = $true)] [string]$PartNumber,
  [switch]$AllowMissingDraftFiles
)

if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
  $staArgs = @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', $MyInvocation.MyCommand.Path,
    '-PartNumber', $PartNumber
  )
  if ($AllowMissingDraftFiles) { $staArgs += '-AllowMissingDraftFiles' }
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
if ([string]::IsNullOrWhiteSpace($pn)) { throw 'Part number is required.' }

function ConvertTo-HtmlMailBody {
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

function Find-NewestNamedZip {
  param(
    [string]$Folder,
    [string]$NameRegex,
    [string]$PartToken
  )
  if ([string]::IsNullOrWhiteSpace($Folder) -or -not (Test-Path -LiteralPath $Folder -PathType Container)) {
    return $null
  }
  $candidates = @(Get-ChildItem -LiteralPath $Folder -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Extension -ieq '.zip' -and $_.Name -match $NameRegex })
  if ($PartToken) {
    $forPart = @($candidates | Where-Object { $_.Name -match [regex]::Escape($PartToken) })
    if ($forPart.Count -gt 0) { $candidates = $forPart }
  }
  if ($candidates.Count -eq 0) { return $null }
  return ($candidates | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1)
}

function Get-SchematicPdfBesidePrjPcb {
  param([string]$SearchRoot)
  $script:SchematicPdfLookupError = $null
  if ([string]::IsNullOrWhiteSpace($SearchRoot) -or -not (Test-Path -LiteralPath $SearchRoot -PathType Container)) {
    $script:SchematicPdfLookupError = "Design folder not found for schematic PDF: $SearchRoot"
    return $null
  }

  $allPrj = @(Get-ChildItem -LiteralPath $SearchRoot -Recurse -File -Filter '*.PrjPcb' -ErrorAction SilentlyContinue |
    Where-Object {
      -not (($_.FullName -split '[\\/]') -contains 'Archive')
    })
  if ($allPrj.Count -eq 0) {
    $script:SchematicPdfLookupError = "No .PrjPcb file found under `"$SearchRoot`" (paths under an `"Archive`" folder are ignored)."
    return $null
  }
  $prj = $allPrj | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
  $prjDir = $prj.DirectoryName

  $beside = @(Get-ChildItem -LiteralPath $prjDir -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Extension -ieq '.pdf' -and $_.BaseName -ieq $prj.BaseName })
  if ($beside.Count -gt 0) {
    return ($beside | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1).FullName
  }

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

  $script:SchematicPdfLookupError = "Schematic PDF not found beside `"$($prj.BaseName).pdf`" or under a PDF folder titled 'Schematics of …'."
  return $null
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

function New-OutlookEmlFile {
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

  foreach ($att in @($AttachmentPaths)) {
    if ([string]::IsNullOrWhiteSpace($att)) { continue }
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
  $MailItem.BodyFormat = 2
  $null = $MailItem.GetInspector()
  Start-Sleep -Milliseconds 400
  $MailItem.HTMLBody = $HtmlBody
}

function Open-OutlookHtmlDraft {
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
      foreach ($att in @($AttachmentPaths)) {
        if (-not [string]::IsNullOrWhiteSpace($att)) {
          [void]$mail.Attachments.Add($att)
        }
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

  $emlPath = Join-Path ([IO.Path]::GetTempPath()) ("wdts-drafts-neel-$([Guid]::NewGuid().ToString('N')).eml")
  New-OutlookEmlFile -To $To -Cc $Cc -Subject $Subject -HtmlBody $HtmlBody -AttachmentPaths $AttachmentPaths -OutputPath $emlPath
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
$partForEmail = $resolved.canonical
if ([string]::IsNullOrWhiteSpace($partForEmail)) { $partForEmail = $pn }
$pcbPart = $resolved.pcbPartNumber
$designDir = $resolved.targetDir
if ([string]::IsNullOrWhiteSpace($designDir)) { $designDir = $resolved.partDir }

Write-Host "  Production document folder: $prodDocDir"
Write-Host "  Part number (email):        $partForEmail"

$fabDraft = Find-NewestNamedZip -Folder $prodDocDir `
  -NameRegex '(?i)^Fabrication\s+.+\s+Rev\s+.+_DRAFT\.zip$' `
  -PartToken $pcbPart
$ppDraft = Find-NewestNamedZip -Folder $prodDocDir `
  -NameRegex '(?i)^Pick Place\s+.+\s+Rev\s+.+_DRAFT\.zip$' `
  -PartToken $partForEmail

if (-not $fabDraft -and -not $ppDraft -and -not $AllowMissingDraftFiles) {
  $message = 'No _DRAFT fabrication or Pick & Place zip was found. Send the email anyway?'
  Write-Host "WARNING: $message"
  Write-Host '---MISSING-DRAFT-FILES-PROMPT---'
  [pscustomobject]@{
    missingDraftFiles = $true
    message           = $message
  } | ConvertTo-Json -Compress
  exit 14
}

$attachmentPaths = New-Object System.Collections.Generic.List[string]
if ($fabDraft) {
  Write-Host "  Fabrication draft zip:      $($fabDraft.FullName)"
  [void]$attachmentPaths.Add($fabDraft.FullName)
}
else {
  Write-Host 'WARNING: Fabrication _DRAFT zip was not found. It will not be attached.'
}
if ($ppDraft) {
  Write-Host "  Pick & Place draft zip:     $($ppDraft.FullName)"
  [void]$attachmentPaths.Add($ppDraft.FullName)
}
else {
  Write-Host 'WARNING: Pick & Place _DRAFT zip was not found. It will not be attached.'
}

$schematicPdf = $null
if (-not [string]::IsNullOrWhiteSpace($designDir)) {
  $schematicPdf = Get-SchematicPdfBesidePrjPcb -SearchRoot $designDir
}
if ($schematicPdf) {
  Write-Host "  Schematic pdf:              $schematicPdf"
  [void]$attachmentPaths.Add($schematicPdf)
}
else {
  $schematicReason = if ($script:SchematicPdfLookupError) { $script:SchematicPdfLookupError } else { 'Schematic PDF not found.' }
  Write-Host "WARNING: $schematicReason The email will be sent without a schematic."
}

$subject = "The $partForEmail files are ready for your review"
$bodyLines = @(
  'Hi Neel,',
  '',
  "The $partForEmail files are ready for your review. They have been uploaded to Arena, but please also find the same files attached for your convenience.",
  'Note P&P files may contain TBD component values at this stage.',
  'Please let me know when your review is complete so we can order the PCBs ASAP.',
  '',
  'Thanks,',
  'Marshall',
  'Email Generated with Script'
)
$htmlBody = ConvertTo-HtmlMailBody -Lines $bodyLines

Write-Host ''
Write-Host "  To:                         nbhattacharya@wdtablesystems.com"
Write-Host "  Cc:                         sgraham@wdtablesystems.com; eosullivan@wdtablesystems.com"
Write-Host "  Subject:                    $subject"
Write-Host ("  Attachments:                {0}" -f $attachmentPaths.Count)

Open-OutlookHtmlDraft `
  -To 'nbhattacharya@wdtablesystems.com' `
  -Cc 'sgraham@wdtablesystems.com; eosullivan@wdtablesystems.com' `
  -Subject $subject `
  -HtmlBody $htmlBody `
  -AttachmentPaths @($attachmentPaths)

exit 0
