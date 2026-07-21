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
