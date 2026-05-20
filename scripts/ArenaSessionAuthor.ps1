# Dot-source from Arena upload/replace scripts.
# Resolves a stable display string for the interactive Windows account running this session.

function Get-ArenaWindowsSessionAuthorName {
  $u = [Environment]::UserName
  $d = [Environment]::UserDomainName
  if ([string]::IsNullOrWhiteSpace($u)) {
    try {
      $ident = [System.Security.Principal.WindowsIdentity]::GetCurrent()
      if ($ident.Name) { return $ident.Name }
    } catch {}
    return 'Unknown user'
  }

  foreach ($path in @(
      $(if ($d) { "WinNT://$d/$u,user" }),
      "WinNT://./$u,user"
    )) {
    if (-not $path) { continue }
    try {
      $dir = New-Object System.DirectoryServices.DirectoryEntry($path)
      $fn = $dir.Properties['fullname'].Value
      if ($null -ne $fn -and "$fn".Trim()) { return "$fn".Trim() }
    } catch { continue }
  }

  try {
    $ident = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    if ($ident.Name -and ($ident.Name -notmatch '^NT AUTHORITY\\')) {
      return $ident.Name
    }
  } catch {}

  if ($d -and $u) { return "$d\$u" }
  return $u
}
