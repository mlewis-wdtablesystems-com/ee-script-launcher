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
  AltiumRevision = revision from folder/PrjPcb (Resolve-RevLetter).
  RevisionOverride = optional user entry; letter must match Altium; zip uses override display.
  #>
  param(
    [Parameter(Mandatory = $true)] [string]$AltiumRevision,
    [string]$RevisionOverride = ''
  )
  $alt = Parse-AltiumRevisionToken -RevisionText $AltiumRevision
  if ([string]::IsNullOrWhiteSpace($RevisionOverride)) {
    return $alt.Display
  }
  $user = Parse-AltiumRevisionToken -RevisionText $RevisionOverride
  if ($alt.Letter -cne $user.Letter) {
    throw (
      "Pick & Place revision letter mismatch: Altium outputs use letter '$($alt.Letter)' " +
      "(from '$AltiumRevision'), but you specified letter '$($user.Letter)'."
    )
  }
  return $user.Display
}

function Get-RevisionLetterOnly {
  param([Parameter(Mandatory = $true)] [string]$RevisionText)
  return (Parse-AltiumRevisionToken -RevisionText $RevisionText).Letter
}

function ConvertTo-PcbaAndBareBoardPartNumbers {
  param([Parameter(Mandatory = $true)] [string]$PartNumber)
  $m = [regex]::Match($PartNumber.Trim(), '^(?<prefix>\d+)-(?<asm>10|15)-(?<suffix>\d+)$')
  if (-not $m.Success) {
    throw "Part number must be a PCBA (PPP-10-XXX or PPP-15-XXX). Got: '$PartNumber'."
  }
  $bareCode = if ($m.Groups['asm'].Value -eq '10') { '20' } else { '25' }
  $pcba = "$($m.Groups['prefix'].Value)-$($m.Groups['asm'].Value)-$($m.Groups['suffix'].Value)"
  $bare = "$($m.Groups['prefix'].Value)-$bareCode-$($m.Groups['suffix'].Value)"
  return [pscustomobject]@{ Pcba = $pcba; Bare = $bare }
}
