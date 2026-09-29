# Dot-source from Open-AltiumPartFolder.ps1, Sync-ArenaPickPlaceZip.ps1

function Parse-AltiumRevisionToken {
  <#
  Parses revision text like D, D1, or "D 1" into letter + optional number.
  Returns Display = letter+number for zip naming (e.g. D1, or D when no number).
  #>
  param([Parameter(Mandatory = $true)] [string]$RevisionText)
  $t = ("$RevisionText").Trim() -replace '\s+', ''
  if (-not $t) { throw 'Revision is empty.' }
  if ($t -notmatch '^([A-Za-z]+)(\d*)$') {
    throw (
      "Invalid revision '$RevisionText'. Use a letter and optional number (e.g. D1 or D 1)."
    )
  }
  $letter = $matches[1].ToUpperInvariant()
  $number = $matches[2]
  $display = if ($number) { "$letter$number" } else { $letter }
  return [pscustomobject]@{
    Letter  = $letter
    Number  = $number
    Display = $display
  }
}

function Resolve-PickPlaceZipRevisionLabel {
  <#
  AltiumRevision = revision from folder/PrjPcb when detectable (optional if user override is set).
  RevisionOverride = user entry (letter + number); used for the zip name. When set, letter mismatch
  against Altium is enforced only when AltiumRevision is non-empty.
  #>
  param(
    [string]$AltiumRevision = '',
    [string]$RevisionOverride = ''
  )
  $overrideTrim = ("$RevisionOverride").Trim()
  if ($overrideTrim) {
    $user = Parse-AltiumRevisionToken -RevisionText $overrideTrim
    $altTrim = ("$AltiumRevision").Trim()
    if ($altTrim) {
      $alt = Parse-AltiumRevisionToken -RevisionText $altTrim
      if ($alt.Letter -cne $user.Letter) {
        throw (
          "Pick & Place revision letter mismatch: Altium outputs use letter '$($alt.Letter)' " +
          "(from '$AltiumRevision'), but you specified letter '$($user.Letter)'."
        )
      }
    }
    return $user.Display
  }
  $altTrim = ("$AltiumRevision").Trim()
  if (-not $altTrim) {
    throw (
      'Could not determine revision letter from the Altium folder or project. ' +
      'Enter Revision (letter + number) in the launcher (e.g. D1).'
    )
  }
  return (Parse-AltiumRevisionToken -RevisionText $altTrim).Display
}

function Get-RevisionLetterOnly {
  param([Parameter(Mandatory = $true)] [string]$RevisionText)
  return (Parse-AltiumRevisionToken -RevisionText $RevisionText).Letter
}

function ConvertTo-PcbaAndBareBoardPartNumbers {
  param([Parameter(Mandatory = $true)] [string]$PartNumber)
  $v = Get-BareBoardPartNumberVariants -PartNumber $PartNumber
  return [pscustomobject]@{ Pcba = $v.Pcba; Bare = $v.Standard }
}

function Get-BareBoardPartNumberVariants {
  param([Parameter(Mandatory = $true)] [string]$PartNumber)
  $m = [regex]::Match($PartNumber.Trim(), '^(?<prefix>\d+)-(?<asm>10|15)-(?<suffix>\d+)$')
  if (-not $m.Success) {
    throw "Part number must be a PCBA (PPP-10-XXX or PPP-15-XXX). Got: '$PartNumber'."
  }
  $p = $m.Groups['prefix'].Value
  $a = $m.Groups['asm'].Value
  $s = $m.Groups['suffix'].Value
  $bareCode = if ($a -eq '10') { '20' } else { '25' }
  return [pscustomobject]@{
    Pcba     = "$p-$a-$s"
    Standard = "$p-$bareCode-$s"
    Legacy   = "$p-$a-$s-PCB"
  }
}

function Get-BomLinePartNumber {
  param($Line)
  if ($null -eq $Line) { return '' }
  if ($Line.item -and $null -ne $Line.item.number) {
    return ("$($Line.item.number)").Trim()
  }
  if ($null -ne $Line.PartNumber) { return ("$($Line.PartNumber)").Trim() }
  if ($null -ne $Line.number) { return ("$($Line.number)").Trim() }
  return ''
}

function Get-BomLineItemName {
  param($Line)
  if ($null -eq $Line) { return '' }
  if ($Line.item -and $null -ne $Line.item.name) {
    return ("$($Line.item.name)").Trim()
  }
  if ($null -ne $Line.Name) { return ("$($Line.Name)").Trim() }
  if ($null -ne $Line.name) { return ("$($Line.name)").Trim() }
  return ''
}

function Test-IsBarePcbBomLine {
  param(
    $Line,
    [string]$PcbaPartNumber = ''
  )
  $num = Get-BomLinePartNumber -Line $Line
  if (-not $num) { return $false }
  if ($PcbaPartNumber -and $num.Equals($PcbaPartNumber.Trim(), [StringComparison]::OrdinalIgnoreCase)) {
    return $false
  }
  if ($num -match '^\d+-(20|25)-\d+$') { return $true }
  if ($num -match '(?i)-PCB$') { return $true }
  $name = Get-BomLineItemName -Line $Line
  if ($name -match '(?i)\bPCB\s*assembl|\bPCBA\b') { return $false }
  if ($name -match '(?i)\bPCB\b|\bbare\s*board\b') { return $true }
  return $false
}

function Get-BarePcbPartNumbersFromBom {
  param(
    [object[]]$BomLines,
    [string]$PcbaPartNumber = ''
  )
  $seen = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
  )
  foreach ($line in @($BomLines)) {
    if (-not (Test-IsBarePcbBomLine -Line $line -PcbaPartNumber $PcbaPartNumber)) { continue }
    $n = Get-BomLinePartNumber -Line $line
    if ($n) { [void]$seen.Add($n) }
  }
  return @($seen | Sort-Object)
}

function Resolve-BareBoardPartNumberFromBom {
  <#
  Prefer the standard PPP-20/25 PCB when it is on the PCBA BOM.
  If it is missing but another PCB line is present, use that PCB for checks.
  #>
  param(
    [Parameter(Mandatory = $true)] [string]$PcbaPartNumber,
    [object[]]$BomLines
  )
  $v = Get-BareBoardPartNumberVariants -PartNumber $PcbaPartNumber
  $standard = $v.Standard
  $legacy = $v.Legacy
  $pcbOnBom = @(Get-BarePcbPartNumbersFromBom -BomLines $BomLines -PcbaPartNumber $v.Pcba)

  $standardOnBom = [bool]($pcbOnBom | Where-Object {
    $_.Equals($standard, [StringComparison]::OrdinalIgnoreCase)
  })
  if ($standardOnBom) {
    return [pscustomobject]@{
      Number          = $standard
      Standard        = $standard
      Source          = 'standard'
      Note            = ''
      OtherPcbNumbers = @($pcbOnBom | Where-Object {
        -not $_.Equals($standard, [StringComparison]::OrdinalIgnoreCase)
      })
    }
  }

  $alternates = @($pcbOnBom | Where-Object {
    -not $_.Equals($standard, [StringComparison]::OrdinalIgnoreCase)
  })

  $chosen = $null
  $source = 'standard-not-in-bom'
  $note = ''
  if ($alternates.Count -gt 0) {
    $legacyOnBom = [bool]($alternates | Where-Object {
      $_.Equals($legacy, [StringComparison]::OrdinalIgnoreCase)
    })
    if ($legacyOnBom) {
      $chosen = $legacy
    } else {
      $coded = @($alternates | Where-Object { $_ -match '^\d+-(20|25)-\d+$' } | Sort-Object)
      if ($coded.Count -gt 0) {
        $chosen = $coded[0]
      } else {
        $chosen = (@($alternates | Sort-Object))[0]
      }
    }
    $source = 'bom-alternate'
    $others = @($alternates | Where-Object {
      -not $_.Equals($chosen, [StringComparison]::OrdinalIgnoreCase)
    })
    $note = "Standard PCB $standard is not in the BOM; using $chosen from the BOM as the PCB for this design."
    if ($others.Count -gt 0) {
      $note += (" Additional PCB part(s) on BOM: {0}." -f ($others -join ', '))
    }
    return [pscustomobject]@{
      Number          = $chosen
      Standard        = $standard
      Source          = $source
      Note            = $note
      OtherPcbNumbers = $others
    }
  }

  return [pscustomobject]@{
    Number          = $standard
    Standard        = $standard
    Source          = $source
    Note            = ''
    OtherPcbNumbers = @()
  }
}
