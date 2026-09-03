# Dot-source from Compare-ArenaAltiumBom.ps1 and Check-AndUploadAltiumBomExport.ps1
# Shared "spec-aware" description reconciliation for Arena item names vs Altium BOM text.

# Optional: canonical PPP-XX-XXX part numbers that skip description reconciliation (known legacy pairs).
$script:BomArenaAltiumDescriptionBypassPartNumbers = [System.Collections.Generic.HashSet[string]]::new(
  [System.StringComparer]::OrdinalIgnoreCase
)
foreach ($x in @(
    '034-10-007'
    # 470 µF polymer cap: Arena ESR (17 mOhm) vs Altium omits ESR token
    '098-13-113'
    # Adjustable buck: Arena "0V" output wording vs Altium "18V" input/rating wording
    '098-10-574'
    # 16-pos receptacle: Arena catalog wording vs Altium "16 way micro-match" wording
    '098-10-551'
    # 8.2 nF cap: Name "8.2nF,25V" vs description "8200 pF" (same value, different tokenization)
    '098-13-167'
    # 10 nF cap: Name "10nF" vs description "10000PF" (10nf vs 10000pf tokens)
    '098-13-038'
    # 3.9 ohm: Name "3R9" vs description "3.9 OHM" (3r9 vs 3.9ohm tokens)
    '098-12-006'
    # FPC connector: Arena "1.00mm" vs Altium "1mm" pitch wording
    '098-10-516'
    # Toroid Transformer (P&P) vs COMMOM MODE CHOKE vertical (Arena) — e.g. 266-10-008 T1
    '031-10-060'
    # 3.3 uH inductor: P&P column shows literal "Description" vs Arena full text — e.g. 190-15-001 L8
    '098-10-584'
    # Terminal block: P&P "Connector" vs Arena full header description — e.g. 190-15-001 J1/J2
    '098-10-508'
    # 2.2 uH wirewound inductor: P&P "IND 2.2UH ... 0.073OHM" vs Arena "2.2 uH ... 73mOhm" — e.g. 272-10-001 L1
    '098-10-461'
  )) {
  [void]$script:BomArenaAltiumDescriptionBypassPartNumbers.Add($x)
}

function Test-BomPartNumberDescriptionCheckBypassed([string]$pn) {
  if ([string]::IsNullOrWhiteSpace($pn)) { return $false }
  $key = if ($pn -match '(?i)(\d{3}-\d{2}-\d{3})') { $matches[1] } else { $pn.Trim() }
  return $script:BomArenaAltiumDescriptionBypassPartNumbers.Contains($key)
}

function Test-BomDescriptionPairCheckBypassed([string]$arenaDesc, [string]$altiumDesc) {
  # Known legacy Arena vs Altium/P&P wording pairs (no shared part-number bypass).
  $a = Normalize-BomDescriptionText $arenaDesc
  $b = Normalize-BomDescriptionText $altiumDesc
  if (-not $a -or -not $b) { return $false }

  # MAA Daughter Board PCB Assembly (P&P) vs AT-MUX DAUGHTER BOARD WITH STANDARD SIZE HEADER PINS (Arena)
  $aIsMaa = [bool]($a -match '(?i)\bmaa\b' -and $a -match '(?i)daughter\s+board')
  $bIsMaa = [bool]($b -match '(?i)\bmaa\b' -and $b -match '(?i)daughter\s+board')
  $aIsAtMux = [bool]($a -match '(?i)\bat[\s\-]?mux\b' -and $a -match '(?i)daughter\s+board')
  $bIsAtMux = [bool]($b -match '(?i)\bat[\s\-]?mux\b' -and $b -match '(?i)daughter\s+board')
  if (($aIsMaa -and $bIsAtMux) -or ($bIsMaa -and $aIsAtMux)) { return $true }

  # CON BNC RCPT STR 18.3MM PCB (P&P) vs BNC Connector Jack, Female Socket 50 Ohms... (Arena)
  $aIsConBnc = [bool]($a -match '(?i)\bcon\b' -and $a -match '(?i)\bbnc\b' -and $a -match '(?i)\brcpt\b')
  $bIsConBnc = [bool]($b -match '(?i)\bcon\b' -and $b -match '(?i)\bbnc\b' -and $b -match '(?i)\brcpt\b')
  $aIsBncJack = [bool]($a -match '(?i)\bbnc\b' -and $a -match '(?i)\bconnector\b' -and $a -match '(?i)\bjack\b')
  $bIsBncJack = [bool]($b -match '(?i)\bbnc\b' -and $b -match '(?i)\bconnector\b' -and $b -match '(?i)\bjack\b')
  if (($aIsConBnc -and $bIsBncJack) -or ($bIsConBnc -and $aIsBncJack)) { return $true }

  return $false
}

function Get-VoltageTokens([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return @() }
  $matches = [regex]::Matches($s, '(?i)(\d+(?:\.\d+)?)\s*v(?:dc|ac)?\b')
  $set = New-Object System.Collections.Generic.HashSet[string]
  foreach ($m in $matches) { [void]$set.Add($m.Groups[1].Value.ToLowerInvariant()) }
  return @($set | Sort-Object)
}

function Normalize-BomDescriptionText([string]$s) {
  if ($null -eq $s) { return '' }
  $t = $s.Trim().ToLowerInvariant()
  $t = $t -replace [char]0x00B5, 'u'   # micro sign
  $t = $t -replace [char]0x03BC, 'u' # Greek mu
  $t = $t -replace '\s+', ' '
  $t = $t -replace '[\u2013\u2014\u2212]', '-' # en/em/minus -> hyphen
  $t = $t -replace '\s*%\s*', '% '
  return $t.Trim()
}

function Get-CurrentTokensFromDesc([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return @() }
  $rx = [regex]'(?i)(\d+(?:\.\d+)?)\s*(ma|ua|µa|a)\b'
  $set = New-Object System.Collections.Generic.HashSet[string]
  foreach ($m in $rx.Matches($s)) {
    $n = $m.Groups[1].Value.ToLowerInvariant()
    $u = $m.Groups[2].Value.ToLowerInvariant()
    if ($u -eq 'ua' -or $u -eq 'µa') { $u = 'ua' }
    [void]$set.Add("$n$u")
  }
  return @($set | Sort-Object)
}

function Convert-BomCurrentTokenToMilliAmps([string]$token) {
  if ([string]::IsNullOrWhiteSpace($token)) { return $null }
  $t = $token.Trim()
  # Order matters: match 'ma' before 'a' so "350ma" is not parsed as ending in 'a'.
  if ($t -match '(?i)^(\d+(?:\.\d+)?)ma$') { return [double]$matches[1] }
  if ($t -match '(?i)^(\d+(?:\.\d+)?)(?:µ|u)a$') { return [double]$matches[1] / 1000.0 }
  if ($t -match '(?i)^(\d+(?:\.\d+)?)a$') { return [double]$matches[1] * 1000.0 }
  return $null
}

function Get-NormalizedCurrentMilliAmpKeysFromDesc([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return @() }
  $keys = New-Object System.Collections.Generic.HashSet[string]
  foreach ($tok in @(Get-CurrentTokensFromDesc $s)) {
    $ma = Convert-BomCurrentTokenToMilliAmps $tok
    if ($null -ne $ma) {
      [void]$keys.Add([string][Math]::Round([double]$ma, 6))
    }
  }
  return @($keys | Sort-Object)
}

function Get-FootprintTokensFromDesc([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return @() }
  $rx = [regex]'(?i)\b(0201|0402|0603|0805|1206|1210|1812|2220|2512|01005|008004|0\d{3})\b'
  $set = New-Object System.Collections.Generic.HashSet[string]
  foreach ($m in $rx.Matches($s)) {
    [void]$set.Add($m.Groups[1].Value.ToUpperInvariant())
  }
  return @($set | Sort-Object)
}

function Get-ValueTokensFromDesc([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return @() }
  # Reuse full description normalization so µ/μ homoglyphs and spacing match other checks.
  $s = Normalize-BomDescriptionText $s
  if ([string]::IsNullOrWhiteSpace($s)) { return @() }
  $set = New-Object System.Collections.Generic.HashSet[string]
  # Capacitors: 4.7uF, 100nF, 10 pF
  foreach ($m in [regex]::Matches($s, '(?i)(\d+(?:\.\d+)?)\s*([pPnNuUmM])?\s*f\b')) {
    [void]$set.Add(("{0}{1}f" -f $m.Groups[1].Value, ($m.Groups[2].Value + '').ToLowerInvariant()).Trim())
  }
  # Resistors: require real "… ohm(s)" words — avoid matching "KOHM" inside impedance ratings.
  foreach ($m in [regex]::Matches($s, '(?i)(\d+(?:\.\d+)?)\s*([kKmMgG])\s*ohms?\b')) {
    [void]$set.Add(("{0}{1}ohm" -f $m.Groups[1].Value, $m.Groups[2].Value.ToLowerInvariant()))
  }
  foreach ($m in [regex]::Matches($s, '(?i)\b(\d+(?:\.\d+)?)\s+ohms?\b')) {
    [void]$set.Add(("{0}ohm" -f $m.Groups[1].Value))
  }
  foreach ($m in [regex]::Matches($s, '(?i)\b(\d+(?:\.\d+)?)\s*([munp])ohms?\b')) {
    [void]$set.Add(("{0}{1}ohm" -f $m.Groups[1].Value, $m.Groups[2].Value.ToLowerInvariant()))
  }
  foreach ($m in [regex]::Matches($s, '(?i)\b(\d+)\s*([rR])\s*(\d+)\b')) {
    [void]$set.Add(('{0}r{1}' -f $m.Groups[1].Value, $m.Groups[3].Value).ToLowerInvariant())
  }
  # Resistor "100R" style (ohms) — common in BOM Name columns.
  foreach ($m in [regex]::Matches($s, '(?i)(?<![a-z0-9])(\d+(?:\.\d+)?)r\b')) {
    [void]$set.Add(("{0}ohm" -f $m.Groups[1].Value))
  }
  # Inductors: 10uH, 4.7 uH
  foreach ($m in [regex]::Matches($s, '(?i)(\d+(?:\.\d+)?)\s*([pPnNuUmM])?\s*h\b')) {
    [void]$set.Add(("{0}{1}h" -f $m.Groups[1].Value, ($m.Groups[2].Value + '').ToLowerInvariant()).Trim())
  }
  # Resistor shorthand without "ohm" (BOM Name column): 200k, 10M — normalize to same *ohm token family.
  foreach ($m in [regex]::Matches($s, '(?i)(?<![a-z0-9])(\d+(?:\.\d+)?)\s*([kKmMgG])(?![a-z])\b')) {
    [void]$set.Add(("{0}{1}ohm" -f $m.Groups[1].Value, $m.Groups[2].Value.ToLowerInvariant()))
  }
  return @($set | Sort-Object)
}

function Get-BomPfValueFromToken([string]$Token) {
  if ($Token -match '^(\d+(?:\.\d+)?)pf$') { return [double]$matches[1] }
  return $null
}

function Remove-BomCapTolerancePfTokens([string[]]$Tokens) {
  <#
  Drop small pF tokens that are tolerance bands (e.g. 0.25pf from "±0.25pF") when a larger
  nominal pF value is also present (e.g. 5.6pf). Applies across assemblies, not one part number.
  #>
  $pfTokens = @($Tokens | Where-Object { $_ -match 'pf$' })
  if ($pfTokens.Count -le 1) { return @($Tokens) }
  $pfValues = @($pfTokens | ForEach-Object { Get-BomPfValueFromToken $_ } | Where-Object { $null -ne $_ })
  if ($pfValues.Count -le 1) { return @($Tokens) }
  $nominal = ($pfValues | Measure-Object -Maximum).Maximum
  if ($nominal -le 0) { return @($Tokens) }
  $keep = New-Object System.Collections.Generic.HashSet[string]
  foreach ($t in $Tokens) {
    if ($t -notmatch 'pf$') {
      [void]$keep.Add($t)
      continue
    }
    $v = Get-BomPfValueFromToken $t
    if ($null -eq $v) {
      [void]$keep.Add($t)
      continue
    }
    if ($v -lt $nominal -and $v -le ($nominal * 0.2)) { continue }
    [void]$keep.Add($t)
  }
  return @($keep | Sort-Object)
}

function Convert-BomValueTokenToMilliOhmKey([string]$Token) {
  $t = $Token.ToLowerInvariant()
  if ($t -match '^(\d+(?:\.\d+)?)mohm$') {
    return [string][Math]::Round([double]$matches[1], 6) + 'mohm'
  }
  if ($t -match '^(\d+(?:\.\d+)?)ohm$') {
    $o = [double]$matches[1]
    if ($o -lt 1000) {
      return [string][Math]::Round($o * 1000.0, 6) + 'mohm'
    }
  }
  return $null
}

function Get-BomCompareValueTokens([string]$Description) {
  $raw = @(Get-ValueTokensFromDesc $Description)
  $filtered = @(Remove-BomCapTolerancePfTokens $raw)
  $keys = New-Object System.Collections.Generic.HashSet[string]
  foreach ($t in $filtered) {
    $mohm = Convert-BomValueTokenToMilliOhmKey $t
    if ($mohm) { [void]$keys.Add($mohm) }
    else { [void]$keys.Add($t) }
  }
  return @($keys | Sort-Object)
}

function Test-BomCompareValueTokenSets {
  param(
    [string[]]$valA,
    [string[]]$valB
  )
  if ($valA.Count -eq 0 -or $valB.Count -eq 0) { return $true }
  $valD = Compare-Object -ReferenceObject $valA -DifferenceObject $valB -SyncWindow 0
  return (-not $valD)
}

function Get-DimensionMmTokensFromDesc([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return @() }
  $set = New-Object System.Collections.Generic.HashSet[string]
  foreach ($m in [regex]::Matches($s, '(?i)(\d+(?:\.\d+)?)\s*(?:x\s*)?(mm|mil)\b')) {
    [void]$set.Add(("{0}{1}" -f $m.Groups[1].Value.ToLowerInvariant(), $m.Groups[2].Value.ToLowerInvariant()))
  }
  return @($set | Sort-Object)
}

function Get-PosCountTokensFromDesc([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return @() }
  $set = New-Object System.Collections.Generic.HashSet[string]
  foreach ($m in [regex]::Matches($s, '(?i)(\d+)\s*pos\b')) {
    [void]$set.Add($m.Groups[1].Value)
  }
  return @($set | Sort-Object)
}

function Test-BomTextLooksCmcOrFerriteBead([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return $false }
  return [bool]($s -match '(?i)\bcmc\b|ferrite\s*bead')
}

function Test-BomCmcFerriteLooseCompatible {
  <#
  CMC / ferrite-bead lines: accept when impedance/ohms wording differs in case/spacing
  if at least one comparable spec token overlaps or one side's value tokens are a subset of the other's.
  #>
  param([Parameter(Mandatory)][string]$a, [Parameter(Mandatory)][string]$b)
  $na = Normalize-BomDescriptionText $a
  $nb = Normalize-BomDescriptionText $b
  if ($na.Contains($nb) -or $nb.Contains($na)) { return $true }
  $valA = @(Get-BomCompareValueTokens $a)
  $valB = @(Get-BomCompareValueTokens $b)
  if ($valA.Count -gt 0 -and $valB.Count -gt 0) {
    $missingA = @($valA | Where-Object { $valB -notcontains $_ })
    $missingB = @($valB | Where-Object { $valA -notcontains $_ })
    if ($missingA.Count -eq 0 -or $missingB.Count -eq 0) { return $true }
    $shared = @($valA | Where-Object { $valB -contains $_ })
    if ($shared.Count -gt 0) { return $true }
  }
  $iA = @(Get-NormalizedCurrentMilliAmpKeysFromDesc $a)
  $iB = @(Get-NormalizedCurrentMilliAmpKeysFromDesc $b)
  if ($iA.Count -gt 0 -and $iB.Count -gt 0) {
    $iCommon = @($iA | Where-Object { $iB -contains $_ })
    if ($iCommon.Count -gt 0) { return $true }
  }
  $fA = @(Get-FootprintTokensFromDesc $a)
  $fB = @(Get-FootprintTokensFromDesc $b)
  if ($fA.Count -gt 0 -and $fB.Count -gt 0) {
    $fCommon = @($fA | Where-Object { $fB -contains $_ })
    if ($fCommon.Count -gt 0) { return $true }
  }
  if (Test-BomLooseTextOverlap -na $na -nb $nb) { return $true }
  return $false
}

function Test-BomLooseTextOverlap {
  param([string]$na, [string]$nb)
  if ([string]::IsNullOrWhiteSpace($na) -or [string]::IsNullOrWhiteSpace($nb)) { return $false }
  $stop = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@('the', 'and', 'for', 'with', 'line', 'metric', 'pcb', 'board', 'assy', 'assembly', 'part', 'type', 'size', 'qty', 'quantity'),
    [System.StringComparer]::OrdinalIgnoreCase
  )
  $wordsA = @([regex]::Split($na, '[^a-z0-9]+') | Where-Object { $_.Length -ge 3 -and -not $stop.Contains($_) })
  $wordsB = @([regex]::Split($nb, '[^a-z0-9]+') | Where-Object { $_.Length -ge 3 -and -not $stop.Contains($_) })
  $setA = [System.Collections.Generic.HashSet[string]]::new([string[]]$wordsA, [System.StringComparer]::OrdinalIgnoreCase)
  $setB = [System.Collections.Generic.HashSet[string]]::new([string[]]$wordsB, [System.StringComparer]::OrdinalIgnoreCase)
  $inter = 0
  foreach ($w in $setB) { if ($setA.Contains($w)) { $inter++ } }
  $den = [Math]::Max(1, [Math]::Min($setA.Count, $setB.Count))
  if ($inter -ge 3) { return $true }
  if ($inter -ge 2 -and ($inter / $den) -ge 0.35) { return $true }
  return $false
}

function Test-BomDescriptionsCompatible {
  param(
    [Parameter(Mandatory)][string]$ArenaDescription,
    [Parameter(Mandatory)][string]$AltiumDescription,
    [Parameter()][string]$PartNumber = '',
    [Parameter()][switch]$AllowExtraValueTokensOnAltiumSide
  )
  $a = $ArenaDescription.Trim()
  $b = $AltiumDescription.Trim()
  if ($a -eq '' -and $b -eq '') { return @{ Ok = $true; Message = '' } }
  if ($a -eq '' -or $b -eq '') {
    return @{ Ok = $false; Message = "Missing description on one side (arena='$a' vs altium='$b')." }
  }

  if ($PartNumber -and (Test-BomPartNumberDescriptionCheckBypassed $PartNumber)) {
    return @{ Ok = $true; Message = '' }
  }
  if (Test-BomDescriptionPairCheckBypassed -arenaDesc $a -altiumDesc $b) {
    return @{ Ok = $true; Message = '' }
  }

  $na = Normalize-BomDescriptionText $a
  $nb = Normalize-BomDescriptionText $b
  if ($na -eq $nb) { return @{ Ok = $true; Message = '' } }
  if ($na.Contains($nb) -or $nb.Contains($na)) { return @{ Ok = $true; Message = '' } }

  if ((Test-BomTextLooksCmcOrFerriteBead $a) -or (Test-BomTextLooksCmcOrFerriteBead $b)) {
    if (Test-BomCmcFerriteLooseCompatible -a $a -b $b) {
      return @{ Ok = $true; Message = '' }
    }
  }

  $vA = @(Get-VoltageTokens $a)
  $vB = @(Get-VoltageTokens $b)
  if ($vA.Count -gt 0 -and $vB.Count -gt 0) {
    $vd = Compare-Object -ReferenceObject $vA -DifferenceObject $vB -SyncWindow 0
    if ($vd) {
      return @{ Ok = $false; Message = ("Voltage tokens differ: arena=[{0}] altium=[{1}]" -f ($vA -join ','), ($vB -join ',')) }
    }
  }

  # Compare currents in a unit-agnostic way (e.g. 350 mA vs 0.35 A).
  $iA = @(Get-NormalizedCurrentMilliAmpKeysFromDesc $a)
  $iB = @(Get-NormalizedCurrentMilliAmpKeysFromDesc $b)
  if ($iA.Count -gt 0 -and $iB.Count -gt 0) {
    $id = Compare-Object -ReferenceObject $iA -DifferenceObject $iB -SyncWindow 0
    if ($id) {
      return @{ Ok = $false; Message = ("Current rating tokens differ (mA equivalent): arena=[{0}] altium=[{1}]" -f ($iA -join ','), ($iB -join ',')) }
    }
  }

  $fA = @(Get-FootprintTokensFromDesc $a)
  $fB = @(Get-FootprintTokensFromDesc $b)
  $valA = @(Get-BomCompareValueTokens $a)
  $valB = @(Get-BomCompareValueTokens $b)
  if ($fA.Count -gt 0 -and $fB.Count -gt 0) {
    $fd = Compare-Object -ReferenceObject $fA -DifferenceObject $fB -SyncWindow 0
    if ($fd) {
      return @{ Ok = $false; Message = ("Footprint / case-code tokens differ: arena=[{0}] altium=[{1}]" -f ($fA -join ','), ($fB -join ',')) }
    }
    # Same package — still require matching R/L/C values when both sides state them (e.g. 270pF vs 240pF).
    if ($valA.Count -gt 0 -and $valB.Count -gt 0) {
      if (-not (Test-BomCompareValueTokenSets -valA $valA -valB $valB)) {
        return @{ Ok = $false; Message = ("Value tokens (R/L/C style) differ: arena=[{0}] altium=[{1}]" -f ($valA -join ','), ($valB -join ',')) }
      }
    }
    # Same case size / package — do not nitpick impedance wording when value tokens are absent (e.g. ferrite beads).
    return @{ Ok = $true; Message = '' }
  }
  if ($valA.Count -gt 0 -and $valB.Count -gt 0) {
    if ($AllowExtraValueTokensOnAltiumSide) {
      $missingFromDesc = @($valA | Where-Object { $valB -notcontains $_ })
      if ($missingFromDesc.Count -gt 0) {
        return @{
          Ok      = $false
          Message = (
            "Value tokens from the short side not found on the long side: short=[{0}] missing [{1}] long=[{2}]" -f `
              ($valA -join ','), ($missingFromDesc -join ','), ($valB -join ',')
          )
        }
      }
    }
    else {
      if (-not (Test-BomCompareValueTokenSets -valA $valA -valB $valB)) {
        return @{ Ok = $false; Message = ("Value tokens (R/L/C style) differ: arena=[{0}] altium=[{1}]" -f ($valA -join ','), ($valB -join ',')) }
      }
    }
  }

  $dimA = @(Get-DimensionMmTokensFromDesc $a)
  $dimB = @(Get-DimensionMmTokensFromDesc $b)
  if ($dimA.Count -gt 0 -and $dimB.Count -gt 0) {
    $dimCommon = @($dimA | Where-Object { $dimB -contains $_ })
    if ($dimCommon.Count -gt 0) { return @{ Ok = $true; Message = '' } }
    return @{ Ok = $false; Message = ("Dimension/mm tokens differ: arena=[{0}] altium=[{1}]" -f ($dimA -join ','), ($dimB -join ',')) }
  }

  $posA = @(Get-PosCountTokensFromDesc $a)
  $posB = @(Get-PosCountTokensFromDesc $b)
  if ($posA.Count -gt 0 -and $posB.Count -gt 0) {
    $posCommon = @($posA | Where-Object { $posB -contains $_ })
    if ($posCommon.Count -gt 0) { return @{ Ok = $true; Message = '' } }
    return @{ Ok = $false; Message = ("Connector / pos-count tokens differ: arena=[{0}] altium=[{1}]" -f ($posA -join ','), ($posB -join ',')) }
  }

  if (Test-BomLooseTextOverlap -na $na -nb $nb) {
    return @{ Ok = $true; Message = '' }
  }

  $specCount = $vA.Count + $vB.Count + $iA.Count + $iB.Count + $fA.Count + $fB.Count + $valA.Count + $valB.Count
  if ($specCount -eq 0) {
    return @{
      Ok      = $false
      Message = "Descriptions differ and no comparable voltage/current/footprint/value tokens were found.`n    Arena : $a`n    Altium: $b"
    }
  }

  # Do not approve when tokens exist on only one side (e.g. package "0603" only in the long text while
  # the short Name is "200k") — that used to fall through as Ok=true and hid BOM placeholder mistakes.
  $hasPairedVoltage = ($vA.Count -gt 0 -and $vB.Count -gt 0)
  $hasPairedCurrent = ($iA.Count -gt 0 -and $iB.Count -gt 0)
  $hasPairedFootprint = ($fA.Count -gt 0 -and $fB.Count -gt 0)
  $hasPairedValue = ($valA.Count -gt 0 -and $valB.Count -gt 0)
  if ($hasPairedVoltage -or $hasPairedCurrent -or $hasPairedFootprint -or $hasPairedValue) {
    return @{ Ok = $true; Message = '' }
  }

  return @{
    Ok      = $false
    Message = (
      "Descriptions could not be reconciled: spec tokens appear on only one side (or only partially overlap), " +
      "which is not enough to treat the strings as compatible.`n    Arena : $a`n    Altium: $b"
    )
  }
}

function Test-BomNameMatchesManufacturerPartNumber {
  <#
  Case- and whitespace-tolerant match between Excel BOM "Name" and Arena manufacturer part number.
  #>
  param(
    [Parameter(Mandatory)][string]$NameText,
    [Parameter(Mandatory)][string]$MpnText
  )
  $n = $NameText.Trim()
  $m = $MpnText.Trim()
  if (-not $n -or -not $m) { return $false }
  if ([string]::Equals($n, $m, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
  $n2 = ($n -replace '\s+', ' ').Trim()
  $m2 = ($m -replace '\s+', ' ').Trim()
  return [string]::Equals($n2, $m2, [System.StringComparison]::OrdinalIgnoreCase)
}

function Test-BomExcelNameVersusDescriptionGeneralCompatible {
  <#
  Internal PJM-sheet check: Altium "Name" or "Comment" column (MPN / short name) vs "03-Description General".
  Reuses the same spec-aware rules as Arena vs Altium BOM compare; optional PartNumber uses the same bypass list.
  #>
  param(
    [Parameter(Mandatory)][string]$NameText,
    [Parameter(Mandatory)][string]$DescriptionGeneralText,
    [string]$PartNumber
  )
  if ($PartNumber -and (Test-BomPartNumberDescriptionCheckBypassed $PartNumber)) {
    return @{ Ok = $true; Message = '' }
  }
  $n = $NameText.Trim()
  $d = $DescriptionGeneralText.Trim()
  if ($n -eq '' -and $d -eq '') { return @{ Ok = $true; Message = '' } }
  # Common: description filled, generic "Name" left blank on non-electrical lines.
  if ($n -eq '' -and $d -ne '') { return @{ Ok = $true; Message = '' } }
  if ($d -eq '' -and $n -ne '') {
    return @{ Ok = $false; Message = '03-Description General is empty but Name is not.' }
  }
  # Placeholder still in description while Name already carries a concrete value/part number.
  if ($d -match '(?i)\bTBD\b' -and $n -notmatch '(?i)\bTBD\b') {
    $nameHasValueToken = (@(Get-ValueTokensFromDesc $n).Count -gt 0)
    $nameLooksMpn = ($n.Length -ge 4) -and ($n -match '[A-Za-z]') -and ($n -match '\d')
    if ($nameHasValueToken -or $nameLooksMpn) {
      return @{
        Ok      = $false
        Message = "03-Description General still contains TBD but Name looks like a concrete value or MPN ('$n'); update the description or align the Name field."
      }
    }
  }
  $chk = Test-BomDescriptionsCompatible -ArenaDescription $n -AltiumDescription $d -AllowExtraValueTokensOnAltiumSide
  if ($chk.Ok) { return $chk }
  return @{
    Ok      = $false
    Message = (
      "Name vs 03-Description General disagree under the same spec-aware rules used for Arena checks.`n    " +
      "Detail: $($chk.Message)"
    )
  }
}

function Format-BomDesignatorsForExcelCheckError([string]$cellText) {
  <#
  Returns a short "Designators: …" fragment for error messages (first 3 refs, then ", ..." if more).
  Splits on comma/semicolon like typical BOM exports.
  #>
  if ([string]::IsNullOrWhiteSpace($cellText)) { return '' }
  $parts = @(($cellText -split '[,;]') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
  if ($parts.Count -eq 0) { return '' }
  $head = @($parts | Select-Object -First 3)
  $s = $head -join ', '
  if ($parts.Count -gt 3) { $s += ', ...' }
  return "Designators: $s"
}

function Get-BomExcelCheckPartLinePrefix {
  param(
    [Parameter(Mandatory)][string]$PartNumber,
    [Parameter(Mandatory)][int]$ExcelRow,
    [string]$DesignatorsCellText
  )
  $prefix = "Part $PartNumber (Excel row $ExcelRow)"
  $d = Format-BomDesignatorsForExcelCheckError $DesignatorsCellText
  if ($d) { $prefix += "; $d" }
  return "${prefix}:"
}

function Get-PjmPartRowsFromWorksheet {
  param(
    [Parameter(Mandatory)] $Worksheet
  )
  $usedRange = $Worksheet.UsedRange
  if ($null -eq $usedRange) { throw "Worksheet '$($Worksheet.Name)' has no used range." }
  $rowCount = [int]$usedRange.Rows.Count
  $colCount = [int]$usedRange.Columns.Count
  $values = $usedRange.Value2

  $headerRow = -1
  $colPn = -1
  $colDesc = -1
  $colName = -1
  $colDesig = -1
  $scan = [Math]::Min($rowCount, 45)
  :scanRows for ($r = 1; $r -le $scan; $r++) {
    for ($c = 1; $c -le $colCount; $c++) {
      $cell = ("" + $values[$r, $c]).Trim()
      if ($cell -eq '01-PJM Part_Number' -or $cell -match '(?i)^01-PJM\s*Part[_ ]?Number$') {
        $headerRow = $r
        $colPn = $c
        break scanRows
      }
    }
  }
  if ($headerRow -lt 0) {
    throw "Could not find header cell '01-PJM Part_Number' on sheet '$($Worksheet.Name)'."
  }

  for ($c = 1; $c -le $colCount; $c++) {
    $h = ("" + $values[$headerRow, $c]).Trim()
    if (-not $h) { continue }
    if ($h -match '(?i)^01-PJM\s*Part[_ ]?Number$' -or $h -eq '01-PJM Part_Number') { continue }
    if ($h -eq '03-Description General' -or $h -match '(?i)^03-\s*Description\s*General$') {
      $colDesc = $c
      break
    }
  }
  if ($colDesc -lt 0) {
    for ($c = 1; $c -le $colCount; $c++) {
      $h = ("" + $values[$headerRow, $c]).Trim()
      if (-not $h) { continue }
      if ($h -match '(?i)^01-PJM\s*Part[_ ]?Number$' -or $h -eq '01-PJM Part_Number') { continue }
      if ($h -match '(?i)01-PJM.*Description') {
        $colDesc = $c
        break
      }
    }
  }
  if ($colDesc -lt 0) {
    for ($c = 1; $c -le $colCount; $c++) {
      $h = ("" + $values[$headerRow, $c]).Trim()
      if ($h -imatch '^Description$') { $colDesc = $c; break }
    }
  }
  if ($colDesc -lt 0) {
    throw ("Could not find a description column (expected '03-Description General', '01-PJM ... Description', or 'Description') on PJM header row {0}." -f $headerRow)
  }

  for ($c = 1; $c -le $colCount; $c++) {
    $h = ("" + $values[$headerRow, $c]).Trim()
    if (-not $h) { continue }
    if ($h -imatch '^Name$') {
      $colName = $c
      break
    }
  }
  if ($colName -lt 0) {
    for ($c = 1; $c -le $colCount; $c++) {
      $h = ("" + $values[$headerRow, $c]).Trim()
      if (-not $h) { continue }
      if ($h -imatch '^Comment$') {
        $colName = $c
        break
      }
    }
  }
  if ($colName -lt 0) {
    for ($c = 1; $c -le $colCount; $c++) {
      $h = ("" + $values[$headerRow, $c]).Trim()
      if (-not $h) { continue }
      if ($h -match '(?i)^02-\s*PJM\s*Name$' -or $h -match '(?i)^01-PJM\s*Name$') {
        $colName = $c
        break
      }
    }
  }
  if ($colName -lt 0) {
    for ($c = 1; $c -le $colCount; $c++) {
      $h = ("" + $values[$headerRow, $c]).Trim()
      if (-not $h) { continue }
      if ($h -match '(?i)^02-\s*PJM\s*Comment$' -or $h -match '(?i)^01-PJM\s*Comment$') {
        $colName = $c
        break
      }
    }
  }

  for ($c = 1; $c -le $colCount; $c++) {
    $h = ("" + $values[$headerRow, $c]).Trim()
    if (-not $h) { continue }
    if ($h -imatch '^Designator$') { $colDesig = $c; break }
  }
  if ($colDesig -lt 0) {
    for ($c = 1; $c -le $colCount; $c++) {
      $h = ("" + $values[$headerRow, $c]).Trim()
      if (-not $h) { continue }
      if ($h -imatch '^Designators$') { $colDesig = $c; break }
    }
  }
  if ($colDesig -lt 0) {
    for ($c = 1; $c -le $colCount; $c++) {
      $h = ("" + $values[$headerRow, $c]).Trim()
      if (-not $h) { continue }
      if ($h -imatch '^Reference\s*Designators?$') { $colDesig = $c; break }
    }
  }
  if ($colDesig -lt 0) {
    for ($c = 1; $c -le $colCount; $c++) {
      $h = ("" + $values[$headerRow, $c]).Trim()
      if (-not $h) { continue }
      if ($h -imatch '^Ref\.?\s*Des\.?$') { $colDesig = $c; break }
    }
  }

  $out = New-Object System.Collections.Generic.List[hashtable]
  for ($r = $headerRow + 1; $r -le $rowCount; $r++) {
    $pn = ("" + $values[$r, $colPn]).Trim()
    if (-not $pn) { continue }
    $desc = if ($colDesc -gt 0) { ("" + $values[$r, $colDesc]).Trim() } else { '' }
    $nameVal = if ($colName -gt 0) { ("" + $values[$r, $colName]).Trim() } else { '' }
    $desigVal = if ($colDesig -gt 0) { ("" + $values[$r, $colDesig]).Trim() } else { '' }
    $out.Add(@{ PartNumber = $pn; Description = $desc; Name = $nameVal; Designators = $desigVal; Row = $r })
  }
  return @{
    SheetName       = [string]$Worksheet.Name
    Rows            = @($out)
    HasNameColumn   = ($colName -gt 0)
    Worksheet     = $Worksheet
    ColPartNumber = $colPn
  }
}

function Get-PjmSheetPartDescriptionRows {
  param(
    [Parameter(Mandatory)] $Workbook,
    [Parameter(Mandatory)][string]$SheetSpec
  )
  $sheetIndex = 0
  if ([int]::TryParse($SheetSpec, [ref]$sheetIndex)) {
    if ($sheetIndex -lt 1 -or $sheetIndex -gt $Workbook.Sheets.Count) {
      throw "PJM sheet index $sheetIndex out of range (1..$($Workbook.Sheets.Count))."
    }
    $ws = $Workbook.Sheets.Item($sheetIndex)
  }
  else {
    $ws = $Workbook.Sheets.Item($SheetSpec)
  }
  return Get-PjmPartRowsFromWorksheet -Worksheet $ws
}
