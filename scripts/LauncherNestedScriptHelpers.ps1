# Dot-source from launcher workflow scripts that invoke nested powershell.exe runs.

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
      ForEach-Object { "$_".TrimEnd() } |
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
    Write-Host $text
  }

  if ($code -ne 0) {
    $summary = Get-NestedScriptFailureSummary -OutputText $text -StepLabel $StepLabel -ExitCode $code
    Write-Host ''
    Write-Host "ERROR: $summary" -ForegroundColor Red
    exit $code
  }

  return $text
}
