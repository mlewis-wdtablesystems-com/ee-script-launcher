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

.PARAMETER Force
  Before starting, stop any process already listening on -Port (e.g. a prior
  launcher / uvicorn). Use when reopening the launcher and the port is still in use.
#>
param(
  [int]$Port = 8765,
  [switch]$NoBrowser,
  [switch]$Force
)

$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here

if ($Force) {
  Write-Host "Force: stopping anything listening on port $Port ..." -ForegroundColor Yellow
  $stopped = $false
  try {
    $pids = @(
      Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue |
        Select-Object -ExpandProperty OwningProcess -Unique |
        Where-Object { $_ -and $_ -ne $PID }
    )
    foreach ($procId in $pids) {
      try {
        $p = Get-Process -Id $procId -ErrorAction SilentlyContinue
        if ($p) {
          Write-Host ("  Stopping PID {0} ({1})" -f $procId, $p.ProcessName)
          Stop-Process -Id $procId -Force -ErrorAction Stop
          $stopped = $true
        }
      } catch {}
    }
  } catch {
    Write-Host ("  Get-NetTCPConnection failed: {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow
  }
  if (-not $stopped) {
    # Fallback: parse netstat (works when Get-NetTCPConnection is unavailable).
    $ns = netstat -ano 2>$null | Select-String -Pattern ":$Port\s+.*LISTENING\s+(\d+)$"
    foreach ($m in $ns) {
      $procId = [int]$m.Matches[0].Groups[1].Value
      if ($procId -and $procId -ne $PID) {
        try {
          $p = Get-Process -Id $procId -ErrorAction SilentlyContinue
          if ($p) {
            Write-Host ("  Stopping PID {0} ({1}) [netstat]" -f $procId, $p.ProcessName)
            Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
            $stopped = $true
          }
        } catch {}
      }
    }
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
