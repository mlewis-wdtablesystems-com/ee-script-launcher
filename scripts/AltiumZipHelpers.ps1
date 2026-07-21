# Dot-source from Open-AltiumPartFolder.ps1 for production-folder zips.

function Copy-AltiumFileForZip {
  param(
    [Parameter(Mandatory = $true)] [string]$SourcePath,
    [Parameter(Mandatory = $true)] [string]$DestPath
  )
  $parent = Split-Path -Parent $DestPath
  if (-not (Test-Path -LiteralPath $parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
  }
  try {
    $src = [System.IO.File]::Open(
      $SourcePath,
      [System.IO.FileMode]::Open,
      [System.IO.FileAccess]::Read,
      [System.IO.FileShare]::ReadWrite
    )
    try {
      $dst = [System.IO.File]::Create($DestPath)
      try {
        $src.CopyTo($dst)
      }
      finally {
        $dst.Dispose()
      }
    }
    finally {
      $src.Dispose()
    }
  }
  catch {
    throw (
      "Cannot read file for zip (close Excel, Altium, or other apps using this file, then retry): " +
      "`n  $SourcePath`n$($_.Exception.Message)"
    )
  }
}

function Copy-AltiumDirectoryForZip {
  param(
    [Parameter(Mandatory = $true)] [string]$SourceDir,
    [Parameter(Mandatory = $true)] [string]$DestDir
  )
  New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
  foreach ($item in @(Get-ChildItem -LiteralPath $SourceDir -Force -ErrorAction Stop)) {
    $destItem = Join-Path $DestDir $item.Name
    if ($item.PSIsContainer) {
      Copy-AltiumDirectoryForZip -SourceDir $item.FullName -DestDir $destItem
    }
    else {
      Copy-AltiumFileForZip -SourcePath $item.FullName -DestPath $destItem
    }
  }
}

function New-AltiumFolderZip {
  <#
  Zip a folder for Arena upload. Copies to a temp staging area (shared-read so open
  CSV/TXT in Excel still copy), builds the zip in %TEMP%, then moves to the destination.
  The zip root includes the source folder name (same layout as Compress-Archive -Path $dir).
  #>
  param(
    [Parameter(Mandatory = $true)] [string]$SourceDirectory,
    [Parameter(Mandatory = $true)] [string]$DestinationZipPath
  )
  if (-not (Test-Path -LiteralPath $SourceDirectory -PathType Container)) {
    throw "Source folder not found: $SourceDirectory"
  }

  $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("wdts-zip-" + [Guid]::NewGuid().ToString('N'))
  New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
  $tempZip = Join-Path $tempRoot 'archive.zip'

  try {
    $leaf = Split-Path -Leaf ($SourceDirectory.TrimEnd('\', '/'))
    $stageFolder = Join-Path $tempRoot $leaf
    Write-Host "  Staging files for zip..."
    Copy-AltiumDirectoryForZip -SourceDir $SourceDirectory -DestDir $stageFolder

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    if (Test-Path -LiteralPath $tempZip) {
      Remove-Item -LiteralPath $tempZip -Force
    }
    [System.IO.Compression.ZipFile]::CreateFromDirectory(
      $stageFolder,
      $tempZip,
      [System.IO.Compression.CompressionLevel]::Optimal,
      $true
    )

    $destParent = Split-Path -Parent $DestinationZipPath
    if (-not (Test-Path -LiteralPath $destParent -PathType Container)) {
      throw "Destination folder not found: $destParent"
    }
    if (Test-Path -LiteralPath $DestinationZipPath) {
      try {
        Remove-Item -LiteralPath $DestinationZipPath -Force
      }
      catch {
        throw (
          "Cannot overwrite existing zip (close it in Explorer or another app, then retry): " +
          $DestinationZipPath
        )
      }
    }

    try {
      Move-Item -LiteralPath $tempZip -Destination $DestinationZipPath -Force
    }
    catch {
      Copy-Item -LiteralPath $tempZip -Destination $DestinationZipPath -Force
    }
  }
  finally {
    if (Test-Path -LiteralPath $tempRoot) {
      Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
  }
}
