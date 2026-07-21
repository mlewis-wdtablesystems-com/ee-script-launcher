# Dot-source from scripts that launch Windows Explorer.

if (-not ('ExplorerWin32' -as [Type])) {
  Add-Type -Namespace 'ExplorerWin32' -Name 'User32' -MemberDefinition @'
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(System.IntPtr hwnd);
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool ShowWindowAsync(System.IntPtr hwnd, int nCmdShow);
'@ | Out-Null
}

function Show-AppWindowToFront {
  param([string]$TitleHint)
  if ([string]::IsNullOrWhiteSpace($TitleHint)) { return }
  Start-Sleep -Milliseconds 700
  try {
    $wsh = New-Object -ComObject 'WScript.Shell'
    [void]$wsh.AppActivate($TitleHint)
  } catch {}
}

function Get-ExplorerShellWindowForPath {
  param([Parameter(Mandatory = $true)] [string]$ExplorerPath)

  $target = $ExplorerPath.TrimEnd('\').ToLowerInvariant()
  $matched = $null
  try {
    $shellApp = New-Object -ComObject Shell.Application
    foreach ($w in $shellApp.Windows()) {
      try {
        $url = $w.LocationURL
        if (-not $url) { continue }
        $local = ([Uri]$url).LocalPath.TrimEnd('\').ToLowerInvariant()
        if ($local -eq $target) {
          $matched = $w
        }
      } catch {}
    }
  } catch {}
  return $matched
}

function Show-ExplorerWindowMaximized {
  <#
  Best-effort: foreground and maximize an Explorer window showing ExplorerPath.
  #>
  param(
    [Parameter(Mandatory = $true)] [string]$ExplorerPath,
    [string]$TitleHint = ''
  )

  $dir = $ExplorerPath.Trim()
  if ([string]::IsNullOrWhiteSpace($dir)) { return }
  if (Test-Path -LiteralPath $dir) {
    $dir = (Resolve-Path -LiteralPath $dir).Path
  }

  $matched = $null
  foreach ($delayMs in @(700, 500, 500)) {
    Start-Sleep -Milliseconds $delayMs
    $matched = Get-ExplorerShellWindowForPath -ExplorerPath $dir
    if ($matched) { break }
  }

  if (-not $matched) {
    try {
      $wsh = New-Object -ComObject 'WScript.Shell'
      if ($TitleHint) { [void]$wsh.AppActivate($TitleHint) }
    } catch {}
    return
  }

  try {
    $hwnd = [IntPtr]$matched.HWND
    [void][ExplorerWin32.User32]::ShowWindowAsync($hwnd, 3)  # SW_MAXIMIZE
    [void][ExplorerWin32.User32]::SetForegroundWindow($hwnd)
  } catch {}
}

function Start-ExplorerWindow {
  <#
  Open a folder in Explorer, or open Explorer with a file selected (/select).
  The window is brought to the foreground and maximized when possible.
  #>
  param(
    [Parameter(ParameterSetName = 'Folder', Mandatory = $true)] [string]$FolderPath,
    [Parameter(ParameterSetName = 'Select', Mandatory = $true)] [string]$SelectPath
  )

  if ($PSCmdlet.ParameterSetName -eq 'Select') {
    $full = $SelectPath.Trim()
    if ([string]::IsNullOrWhiteSpace($full)) { return }
    if (Test-Path -LiteralPath $full) {
      $full = (Resolve-Path -LiteralPath $full).Path
    }
    Start-Process -FilePath 'explorer.exe' -ArgumentList "/select,`"$full`""
    $dir = Split-Path -Parent $full
    Show-ExplorerWindowMaximized -ExplorerPath $dir -TitleHint (Split-Path -Leaf $dir)
    return
  }

  $dir = $FolderPath.Trim()
  if ([string]::IsNullOrWhiteSpace($dir)) { return }
  if (Test-Path -LiteralPath $dir) {
    $dir = (Resolve-Path -LiteralPath $dir).Path
  }
  Start-Process -FilePath 'explorer.exe' -ArgumentList $dir
  Show-ExplorerWindowMaximized -ExplorerPath $dir -TitleHint (Split-Path -Leaf $dir)
}
