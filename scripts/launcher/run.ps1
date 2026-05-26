<#
.SYNOPSIS
  Bootstrap and start the local script launcher (FastAPI on http://127.0.0.1:8765).

.DESCRIPTION
  - Creates a .venv on first run and installs requirements.txt.
  - Re-installs only when requirements.txt changes (.installed marker tracks the hash).
  - Starts uvicorn (foreground) and opens the browser after the server is up.

  Arena env vars (ARENA_EMAIL, ARENA_PASSWORD, ARENA_WORKSPACE_ID) are inherited
  from the shell that runs this script — set them first if you want the Arena
  cards to work.

.PARAMETER Port
  HTTP port. Default 8765.

.PARAMETER NoBrowser
  Don't open the browser automatically.

.PARAMETER NoForce
  Do not stop an existing launcher already listening on -Port before starting.
  By default, any prior instance on that port is terminated.
#>
param(
  [int]$Port = 8765,
  [switch]$NoBrowser,
  [switch]$NoForce
)

$ErrorActionPreference = 'Stop'

function Stop-ProcessesListeningOnPort {
  param([Parameter(Mandatory = $true)] [int]$PortNum)

  $stopped = $false
  $seen = New-Object System.Collections.Generic.HashSet[int]

  foreach ($procId in @(
    Get-NetTCPConnection -LocalPort $PortNum -State Listen -ErrorAction SilentlyContinue |
      Select-Object -ExpandProperty OwningProcess -Unique
  )) {
    if (-not $procId -or $procId -eq $PID) { continue }
    if ($seen.Add($procId)) {
      try {
        $p = Get-Process -Id $procId -ErrorAction SilentlyContinue
        if ($p) {
          Write-Host ("  Stopping PID {0} ({1})" -f $procId, $p.ProcessName)
          Stop-Process -Id $procId -Force -ErrorAction Stop
          $stopped = $true
        }
      } catch {
        Write-Host ("  Could not stop PID {0}: {1}" -f $procId, $_.Exception.Message) -ForegroundColor DarkYellow
      }
    }
  }

  if (-not $stopped) {
    try {
      $ns = netstat -ano 2>$null | Select-String -Pattern (":$PortNum\s+.*LISTENING\s+(\d+)$")
      foreach ($m in $ns) {
        $procId = [int]$m.Matches[0].Groups[1].Value
        if (-not $procId -or $procId -eq $PID) { continue }
        if (-not $seen.Add($procId)) { continue }
        try {
          $p = Get-Process -Id $procId -ErrorAction SilentlyContinue
          if ($p) {
            Write-Host ("  Stopping PID {0} ({1}) [netstat]" -f $procId, $p.ProcessName)
            Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
            $stopped = $true
          }
        } catch {}
      }
    } catch {
      Write-Host ("  netstat fallback failed: {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow
    }
  }

  return $stopped
}

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

if (-not $NoForce) {
  Write-Host "Stopping any existing launcher on port $Port ..." -ForegroundColor Yellow
  try {
    $null = Stop-ProcessesListeningOnPort -PortNum $Port
  } catch {
    Write-Host ("  Port cleanup warning: {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow
  }
  Start-Sleep -Milliseconds 500
}

$python = (Get-Command python -ErrorAction SilentlyContinue)
if (-not $python) {
  Write-Host 'python is not on PATH. Install Python 3.10+ first.' -ForegroundColor Red
  exit 1
}

$venv = Join-Path $here '.venv'
$venvPython = Join-Path $venv 'Scripts\python.exe'
$requirements = Join-Path $here 'requirements.txt'
$marker = Join-Path $here '.installed'

if (-not (Test-Path -LiteralPath $venvPython)) {
  Write-Host 'Creating virtual environment...' -ForegroundColor Cyan
  & python -m venv $venv
  if ($LASTEXITCODE -ne 0) { throw 'venv creation failed.' }
}

$reqHash = (Get-FileHash -LiteralPath $requirements -Algorithm SHA256).Hash
$installedHash = if (Test-Path -LiteralPath $marker) { Get-Content -LiteralPath $marker -Raw } else { '' }

if ($reqHash.Trim() -ne $installedHash.Trim()) {
  Write-Host 'Installing dependencies...' -ForegroundColor Cyan
  & $venvPython -m pip install --upgrade pip | Out-Null
  & $venvPython -m pip install -r $requirements
  if ($LASTEXITCODE -ne 0) { throw 'pip install failed.' }
  Set-Content -LiteralPath $marker -Value $reqHash -NoNewline
}

if (-not $NoBrowser) {
  Start-Job -ScriptBlock {
    param($url)
    Start-Sleep -Seconds 2
    Start-Process $url
  } -ArgumentList "http://127.0.0.1:$Port" | Out-Null
}

Write-Host "Starting launcher on http://127.0.0.1:$Port  (Ctrl+C to stop)" -ForegroundColor Green
& $venvPython -m uvicorn app:app --host 127.0.0.1 --port $Port
