# Dot-source from launcher workflow scripts that invoke nested powershell.exe runs.
#
# Result colours (console has no DarkOrange; DarkYellow is the warning/orange):
#   good = DarkGreen, bad = DarkRed, warning = DarkYellow

function Get-ScriptResultLineColor {
  param([string]$Line)
  $t = ("$Line").Trim()
  if (-not $t) { return $null }

  if ($t -match '(?i)PCB part\(s\) OK missing from P&P|board item - not required') { return $null }

  if ($t -match '(?i)\(\s*0\s*\)') { return 'DarkGreen' }
  if ($t -match '(?i)\(none\)') { return 'DarkGreen' }
  if ($t -match '(?i)\(no (quantity|description|duplicate)') { return 'DarkGreen' }

  if ($t -match '^=== ') {
    if ($t -match '(?i)fail|issue|error') { return 'DarkRed' }
    if ($t -match '(?i)match|passed|All compared|All required|All checks|OK') { return 'DarkGreen' }
    if ($t -match '(?i)warn') { return 'DarkYellow' }
    return $null
  }

  if ($t -match '(?i)^(WARN:|WARNING:|SKIP:)') { return 'DarkYellow' }
  if ($t -match '(?i)\bWARN(ING)?:') { return 'DarkYellow' }
  if ($t -match '(?i)^NOTE:|assuming Rev|Preview only|Completed with warnings') { return 'DarkYellow' }

  if ($t -match '(?i)^(FAILED:|ERROR:|MISSING:)') { return 'DarkRed' }
  if ($t -match '(?i)\bERROR:|\bFAILED\b|\bMISSING:') { return 'DarkRed' }
  if ($t -match '(?i)Check finished with issues|issue\(s\)|difference\(s\)|letters differ') { return 'DarkRed' }
  if ($t -match '(?i)mismatch') { return 'DarkRed' }
  if ($t -match '(?i)CHANGED |ADDED on working|REMOVED on working|REVISION NUMBER CHANGED') { return 'DarkRed' }
  if ($t -match '(?i)not found in|has no trailing|expected Rev |could not parse|could not load|could not determine') { return 'DarkRed' }
  if ($t -match '(?i)not on Arena BOM|missing from P&P|empty Description|P&P rows for Arena NF') { return 'DarkRed' }
  if ($t -match '(?i)is not in the BOM of PCBA') { return 'DarkRed' }
  if ($t -match '(?i)P&P text:|Pick & Place:|Arena BOM\s*:|Detail\s*:') { return 'DarkRed' }
  if ($t -match '(?i)Exception|CategoryInfo|FullyQualifiedErrorId|^At line |^\+ ') { return 'DarkRed' }
  if ($t -match '(?i)\.ps1\s*:') { return 'DarkRed' }
  if ($t -match '(?i)zip aborted|checks failed|\(exit \d+\)') { return 'DarkRed' }

  if ($t -match '(?i)^OK:|\bOK\b|-> Rev .+ OK') { return 'DarkGreen' }
  if ($t -match '(?i)Revision letters match|All checks passed|All compared|All required|All steps completed') { return 'DarkGreen' }
  if ($t -match '(?i)^Matched ') { return 'DarkGreen' }
  if ($t -match '(?i)Found bare board| -> Found ') { return 'DarkGreen' }
  if ($t -match '(?i)^Done\.?$|Done\. |passed') { return 'DarkGreen' }
  if ($t -match '(?i)no changes detected|lines match|fields match|unchanged') { return 'DarkGreen' }
  if ($t -match '(?i)No Illustrative Image|Compare finished with no') { return 'DarkGreen' }
  if ($t -match '(?i)Accepted N/F') { return 'DarkGreen' }

  return $null
}

function Get-ScriptResultAnsiRgb {
  param([string]$ColorName)
  switch ($ColorName) {
    'DarkGreen'  { return '31;122;69' }
    'DarkRed'    { return '139;21;21' }
    'DarkYellow' { return '192;86;18' }
    default      { return $null }
  }
}

function Write-ScriptResultLines {
  param([string]$Text)
  # Write to the process stdout handle, not the PowerShell pipeline.
  # Write-Output would be captured as extra "results" by callers (e.g. Check Ready
  # for Production Release treating every log line as a failed check).
  $esc = [char]27
  foreach ($line in @($Text -split "`r?`n", -1)) {
    $color = Get-ScriptResultLineColor -Line $line
    $rgb = Get-ScriptResultAnsiRgb -ColorName $color
    $outLine = if ($rgb) {
      "{0}[38;2;{1}m{2}{0}[0m" -f $esc, $rgb, $line
    } else {
      $line
    }
    [Console]::Out.WriteLine($outLine)
  }
}

function Format-NestedPowerShellOutput {
  param([AllowEmptyCollection()][object[]]$Records)
  ($Records | ForEach-Object {
    if ($_ -is [System.Management.Automation.ErrorRecord]) {
      $msg = $_.Exception.Message
      if ($msg) { $msg } else { $_.ToString() }
    }
    else {
      "$_"
    }
  }) -join "`n"
}

function Get-NestedScriptFailureSummary {
  <#
  Pull a short, user-facing failure line from nested script output (stdout + stderr).
  #>
  param(
    [string]$OutputText,
    [string]$StepLabel,
    [int]$ExitCode
  )
  $lines = @(
    $OutputText -split "`r?`n" |
      ForEach-Object { (("$_") -replace '\x1b\[[0-9;]*m', '').TrimEnd() } |
      Where-Object { $_.Trim().Length -gt 0 }
  )

  $errorLines = @($lines | Where-Object { $_ -match '\bERROR:' })
  if ($errorLines.Count -gt 0) {
    return ($errorLines[-1] -replace '^\s*ERROR:\s*', '').Trim()
  }

  $aborted = @($lines | Where-Object { $_ -match 'zip aborted:' })
  if ($aborted.Count -gt 0) {
    return $aborted[-1]
  }

  for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i] -match '^Could not determine revision letter') {
      $msg = $lines[$i]
      for ($j = $i + 1; $j -lt $lines.Count; $j++) {
        if ($lines[$j] -match '^(At |\+ |CategoryInfo|FullyQualifiedErrorId|\s*$)') { break }
        $msg += ' ' + $lines[$j].Trim()
      }
      return ($msg -replace '\s+', ' ').Trim()
    }
  }

  $throwLines = @($lines | Where-Object {
      $_ -match '^(Fabrication zip aborted|Gerber zip aborted|None of the required folders|Production document|PFS folder|No matching folder found|Missing folder:|Fabrication zip not found|Check and Build Fabrication Zip failed|Check and Build Gerber Zip failed)'
    })
  if ($throwLines.Count -gt 0) {
    return $throwLines[-1]
  }

  $psError = @($lines | Where-Object { $_ -match '^Open-AltiumPartFolder\.ps1\s*:\s*(.+)$' })
  if ($psError.Count -gt 0) {
    $msg = ($psError[-1] -replace '^Open-AltiumPartFolder\.ps1\s*:\s*', '').Trim()
    if ($msg) { return $msg }
  }

  if ($lines.Count -gt 0) {
    return $lines[-1]
  }

  return "$StepLabel failed (exit $ExitCode)."
}

function Invoke-NestedLauncherPowerShell {
  <#
  Run powershell.exe, capture merged output without terminating on child stderr,
  echo it to the host, and exit with the child code when non-zero.
  #>
  param(
    [Parameter(Mandatory = $true)][string[]]$ArgumentList,
    [string]$StepLabel = 'Nested script'
  )

  $prev = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'
  try {
    $out = & powershell.exe @ArgumentList 2>&1
    $code = $LASTEXITCODE
  }
  finally {
    $ErrorActionPreference = $prev
  }

  $text = Format-NestedPowerShellOutput $out
  if ($text.Trim()) {
    Write-ScriptResultLines $text
  }

  if ($code -ne 0) {
    $summary = Get-NestedScriptFailureSummary -OutputText $text -StepLabel $StepLabel -ExitCode $code
    Write-Host ''
    Write-Host "ERROR: $summary" -ForegroundColor DarkRed
    exit $code
  }

  return $text
}
