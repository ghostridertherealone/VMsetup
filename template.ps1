The easiest way to randomly rearrange your pinned taskbar shortcuts on Windows 10, without delving into complex and risky direct registry manipulation of binary data, is to programmatically unpin the existing shortcuts and then re-pin them in a random order. This method uses standard shell operations.

This PowerShell script will:

Identify the target applications of your currently pinned taskbar shortcuts (those typically found in the User Pinned\TaskBar folder).
Attempt to unpin all of these items.
Shuffle the list of these applications randomly.
Attempt to re-pin them in the new random order.
The script runs silently without prompts and assumes it has administrative privileges. Due to the nature of unpinning and re-pinning, you will see the icons disappear and reappear on your taskbar during the script's execution. This method primarily works for standard desktop applications that are pinned by the user. Some system-pinned items or Universal Windows Platform (UWP) apps might behave differently.

PowerShell

#Requires -RunAsAdministrator

# --- Configuration ---
# Verb names for context menu actions. These are typically for English Windows installations.
# If your Windows is in a different language, these might need to be adjusted.
# You can find the correct verb by right-clicking an app/shortcut and seeing the exact text for
# "Pin to taskbar" or "Unpin from taskbar". The '&' indicates an accelerator key.
$verbUnpinFromTaskbar = "Unpin from Tas&kbar"
$verbPinToTaskbar = "Pin to Tas&kbar"

# Folder where shortcuts for user-pinned taskbar items are usually stored.
$userPinnedItemsFolder = Join-Path -Path $env:APPDATA -ChildPath "Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"

# Delay in milliseconds after each pin/unpin action to allow the system to process.
$actionDelayMilliseconds = 500

# --- Step 1: Identify target executables of currently pinned items ---
$listOfTargetExecutables = [System.Collections.Generic.List[string]]::new()

if (Test-Path $userPinnedItemsFolder) {
    $windowsShell = New-Object -ComObject WScript.Shell
    # Get a static list of shortcuts at the beginning.
    $currentShortcutsInFolder = Get-ChildItem -Path $userPinnedItemsFolder -Filter *.lnk -ErrorAction SilentlyContinue
    
    foreach ($shortcutFileItem in $currentShortcutsInFolder) {
        try {
            $shortcutObject = $windowsShell.CreateShortcut($shortcutFileItem.FullName)
            # Ensure the target path is valid and the target file exists.
            if ($shortcutObject -and -not [string]::IsNullOrWhiteSpace($shortcutObject.TargetPath) -and (Test-Path $shortcutObject.TargetPath -PathType Leaf -ErrorAction SilentlyContinue)) {
                $listOfTargetExecutables.Add($shortcutObject.TargetPath)
            }
        } catch {
            # Silently ignore errors if a shortcut is problematic.
        }
    }
}

# Remove duplicate target paths, if any.
$uniqueTargetExecutables = $listOfTargetExecutables | Get-Unique

# If there are fewer than 2 items, reordering doesn't make sense.
if ($uniqueTargetExecutables.Count -lt 2) {
    exit
}

# --- Step 2: Unpin all items currently in the User Pinned\TaskBar folder ---
$shellApplicationObject = New-Object -ComObject Shell.Application
$taskbarFolderComObject = $shellApplicationObject.Namespace($userPinnedItemsFolder)

if ($taskbarFolderComObject) {
    # Create a static list of items to attempt to unpin from the COM object's perspective.
    $itemsToAttemptUnpin = @($taskbarFolderComObject.Items())

    foreach ($comItem in $itemsToAttemptUnpin) {
        try {
            $unpinActionVerb = $comItem.Verbs() | Where-Object { $_.Name -eq $verbUnpinFromTaskbar }
            if ($unpinActionVerb) {
                $unpinActionVerb.DoIt()
                Start-Sleep -Milliseconds $actionDelayMilliseconds
            }
        } catch {
            # Silently ignore if unpinning fails for a specific item.
        }
    }
    # Allow a brief moment for all unpin operations to settle.
    Start-Sleep -Seconds 2
}

# --- Step 3: Shuffle the list of target executables ---
$shuffledTargetExecutables = $uniqueTargetExecutables | Get-Random -Count $uniqueTargetExecutables.Count

# --- Step 4: Re-pin the executables in the new random order ---
foreach ($executablePath in $shuffledTargetExecutables) {
    if (Test-Path $executablePath -PathType Leaf) { # Ensure it's a file.
        $executableDirectory = Split-Path -Path $executablePath
        $executableName = Split-Path -Path $executablePath -Leaf
        
        try {
            $directoryComObject = $shellApplicationObject.Namespace($executableDirectory)
            if ($directoryComObject) {
                $fileComObject = $directoryComObject.ParseName($executableName)
                if ($fileComObject) {
                    $pinActionVerb = $fileComObject.Verbs() | Where-Object { $_.Name -eq $verbPinToTaskbar }
                    if ($pinActionVerb) {
                        $pinActionVerb.DoIt()
                        Start-Sleep -Milliseconds $actionDelayMilliseconds
                    }
                }
            }
        } catch {
            # Silently ignore if pinning fails for a specific item.
        }
    }
}
