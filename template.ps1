# Random Desktop Shortcut Rearranger
# Randomly repositions all shortcuts on the desktop

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Drawing;

public class Win32 {
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
    
    [DllImport("user32.dll")]
    public static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);
    
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
    
    public const uint LVM_GETITEMCOUNT = 0x1004;
    public const uint LVM_GETITEMPOSITION = 0x1010;
    public const uint LVM_SETITEMPOSITION = 0x100F;
}
"@

# Get desktop window handles
$progman = [Win32]::FindWindow("Progman", "Program Manager")
$desktopListView = [Win32]::FindWindowEx($progman, [IntPtr]::Zero, "SHELLDLL_DefView", $null)
$desktopListView = [Win32]::FindWindowEx($desktopListView, [IntPtr]::Zero, "SysListView32", $null)

if ($desktopListView -eq [IntPtr]::Zero) {
    # Try alternative method for Windows 10/11
    $workerW = [Win32]::FindWindow("WorkerW", $null)
    while ($workerW -ne [IntPtr]::Zero) {
        $desktopListView = [Win32]::FindWindowEx($workerW, [IntPtr]::Zero, "SHELLDLL_DefView", $null)
        if ($desktopListView -ne [IntPtr]::Zero) {
            $desktopListView = [Win32]::FindWindowEx($desktopListView, [IntPtr]::Zero, "SysListView32", $null)
            break
        }
        $workerW = [Win32]::FindWindowEx([IntPtr]::Zero, $workerW, "WorkerW", $null)
    }
}

if ($desktopListView -eq [IntPtr]::Zero) {
    Write-Host "Could not find desktop ListView. Try running as administrator or check Windows version compatibility." -ForegroundColor Red
    exit 1
}

# Get desktop dimensions
$rect = New-Object Win32+RECT
[Win32]::GetWindowRect($desktopListView, [ref]$rect)
$desktopWidth = $rect.Right - $rect.Left
$desktopHeight = $rect.Bottom - $rect.Top

# Get number of desktop items
$itemCount = [Win32]::SendMessage($desktopListView, [Win32]::LVM_GETITEMCOUNT, 0, 0)

if ($itemCount -eq 0) {
    Write-Host "No desktop shortcuts found to rearrange." -ForegroundColor Yellow
    exit 0
}

Write-Host "Found $itemCount desktop shortcuts. Randomizing positions..." -ForegroundColor Green

# Create random number generator
$random = New-Object System.Random

# Define grid parameters (approximate icon spacing)
$iconWidth = 75
$iconHeight = 75
$marginX = 20
$marginY = 20

# Calculate grid dimensions
$gridCols = [Math]::Floor(($desktopWidth - $marginX * 2) / $iconWidth)
$gridRows = [Math]::Floor(($desktopHeight - $marginY * 2) / $iconHeight)

# Create list of available grid positions
$availablePositions = @()
for ($row = 0; $row -lt $gridRows; $row++) {
    for ($col = 0; $col -lt $gridCols; $col++) {
        $x = $marginX + ($col * $iconWidth)
        $y = $marginY + ($row * $iconHeight)
        $availablePositions += @{ X = $x; Y = $y }
    }
}

# Shuffle available positions
for ($i = $availablePositions.Count - 1; $i -gt 0; $i--) {
    $j = $random.Next(0, $i + 1)
    $temp = $availablePositions[$i]
    $availablePositions[$i] = $availablePositions[$j]
    $availablePositions[$j] = $temp
}

# Reposition each desktop item
for ($i = 0; $i -lt $itemCount; $i++) {
    if ($i -lt $availablePositions.Count) {
        $pos = $availablePositions[$i]
        $lParam = ($pos.Y -shl 16) -bor ($pos.X -band 0xFFFF)
        [Win32]::SendMessage($desktopListView, [Win32]::LVM_SETITEMPOSITION, $i, $lParam)
    }
}

# Refresh desktop
$null = [Win32]::SendMessage($desktopListView, 0x111, 0, 0) # WM_COMMAND with refresh

Write-Host "Desktop shortcuts have been randomly rearranged!" -ForegroundColor Green
Write-Host "Press F5 on desktop or restart Explorer if icons don't update immediately." -ForegroundColor Cyan
