# Get the script directory and Desktop_Shortcuts folder path
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$shortcutsFolder = Join-Path $scriptDir "Desktop_Shortcuts"
$desktopPath = [Environment]::GetFolderPath("Desktop")

# Check if Desktop_Shortcuts folder exists
if (-not (Test-Path $shortcutsFolder)) {
    Write-Error "Desktop_Shortcuts folder not found in script directory: $scriptDir"
    exit 1
}

# Get all shortcut files (.lnk) from the folder
$shortcuts = Get-ChildItem -Path $shortcutsFolder -Filter "*.lnk" -File

# Check if there are any shortcuts
if ($shortcuts.Count -eq 0) {
    Write-Warning "No shortcut files (.lnk) found in Desktop_Shortcuts folder"
    exit 0
}

# Determine how many shortcuts to select (max 5 or total available)
$selectCount = [Math]::Min(5, $shortcuts.Count)

# Randomly select shortcuts
$selectedShortcuts = $shortcuts | Get-Random -Count $selectCount

Write-Host "Copying $($selectedShortcuts.Count) random shortcuts to desktop:"

# Copy each selected shortcut to desktop
foreach ($shortcut in $selectedShortcuts) {
    $destinationPath = Join-Path $desktopPath $shortcut.Name
    
    try {
        Copy-Item -Path $shortcut.FullName -Destination $destinationPath -Force
        Write-Host "✓ Copied: $($shortcut.Name)"
    }
    catch {
        Write-Error "Failed to copy $($shortcut.Name): $($_.Exception.Message)"
    }
}

Write-Host "`nCompleted! $($selectedShortcuts.Count) shortcuts added to desktop."
