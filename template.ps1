Requires -RunAsAdministrator

$scriptRoot = $PSScriptRoot
$customShortcutsFolder = Join-Path -Path $scriptRoot -ChildPath "Taskbar_Shortcuts"

$verbUnpin = "Unpin from Tas&kbar"
$verbPin = "Pin to Tas&kbar"
$actionDelayMilliseconds = 300

$shellApplication = New-Object -ComObject Shell.Application

$standardUserPinnedFolder = Join-Path -Path $env:APPDATA -ChildPath "Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
if (Test-Path $standardUserPinnedFolder) {
    $folderObject = $shellApplication.Namespace($standardUserPinnedFolder)
    if ($folderObject) {
        $itemsToUnpin = @($folderObject.Items())
        foreach ($itemToUnpin in $itemsToUnpin) {
            try {
                $unpinActionVerb = $itemToUnpin.Verbs() | Where-Object { $_.Name -eq $verbUnpin }
                if ($unpinActionVerb) {
                    $unpinActionVerb.DoIt()
                    Start-Sleep -Milliseconds $actionDelayMilliseconds
                }
            } catch {}
        }
        Start-Sleep -Seconds 1
    }
}

if (-not (Test-Path $customShortcutsFolder)) {
    exit
}

$availableShortcutsList = Get-ChildItem -Path $customShortcutsFolder -Filter *.lnk -ErrorAction SilentlyContinue
if ($availableShortcutsList.Count -eq 0) {
    exit
}

$minimumPins = 5
$maximumPins = 7
$numberOfPinsToSelect = Get-Random -Minimum $minimumPins -Maximum ($maximumPins + 1)

if ($numberOfPinsToSelect -gt $availableShortcutsList.Count) {
    $numberOfPinsToSelect = $availableShortcutsList.Count
}

if ($numberOfPinsToSelect -eq 0) {
    exit
}

$selectedShortcutsToPin = $availableShortcutsList | Get-Random -Count $numberOfPinsToSelect

foreach ($shortcutItem in $selectedShortcutsToPin) {
    try {
        $shortcutDirectoryObject = $shellApplication.Namespace($shortcutItem.DirectoryName)
        if ($shortcutDirectoryObject) {
            $shortcutFileObject = $shortcutDirectoryObject.ParseName($shortcutItem.Name)
            if ($shortcutFileObject) {
                $pinActionVerb = $shortcutFileObject.Verbs() | Where-Object { $_.Name -eq $verbPin }
                if ($pinActionVerb) {
                    $pinActionVerb.DoIt()
                    Start-Sleep -Milliseconds $actionDelayMilliseconds
                }
            }
        }
    } catch {}
}
