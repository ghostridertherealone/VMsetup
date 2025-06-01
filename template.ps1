# Random Desktop Shortcut Rearranger - Windows 10 Compatible
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
    
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
    
    [DllImport("user32.dll")]
    public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
    
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    
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
    public const uint WM_COMMAND = 0x111;
}
"@

function Find-DesktopListView {
    $desktopHandle = $null
    
    # Method 1: Traditional approach
    $progman = [Win32]::FindWindow("Progman", "Program Manager")
    if ($progman -ne [IntPtr]::Zero) {
        $defView = [Win32]::FindWindowEx($progman, [IntPtr]::Zero, "SHELLDLL_DefView", $null)
        if ($defView -ne [IntPtr]::Zero) {
            $listView = [Win32]::FindWindowEx($defView, [IntPtr]::Zero, "SysListView32", $null)
            if ($listView -ne [IntPtr]::Zero) {
                return $listView
            }
        }
    }
    
    # Method 2: Send message to Progman to spawn WorkerW
    [Win32]::SendMessage($progman, 0x052C, 0, 0)
    
    # Method 3: Enumerate WorkerW windows
    $callback = {
        param($hwnd, $lParam)
        $className = New-Object System.Text.StringBuilder(256)
        [Win32]::GetClassName($hwnd, $className, $className.Capacity)
        
        if ($className.ToString() -eq "WorkerW") {
            $defView = [Win32]::FindWindowEx($hwnd, [IntPtr]::Zero, "SHELLDLL_DefView", $null)
            if ($defView -ne [IntPtr]::Zero) {
                $listView = [Win32]::FindWindowEx($defView, [IntPtr]::Zero, "SysListView32", $null)
                if ($listView -ne [IntPtr]::Zero) {
                    $script:desktopHandle = $listView
                    return $false # Stop enumeration
                }
            }
        }
        return $true # Continue enumeration
    }
    
    $script:desktopHandle = $null
    [Win32]::EnumWindows($callback, [IntPtr]::Zero)
    
    return $script:desktopHandle
}

# Find desktop ListView
$desktopListView = Find-DesktopListView

if ($desktopListView -eq [IntPtr]::Zero -or $desktopListView -eq $null) {
    Write-Host "Could not find desktop ListView. Trying alternative method..." -ForegroundColor Yellow
    
    # Alternative: Use COM objects to manipulate desktop
    try {
        $shell = New-Object -ComObject Shell.Application
        $desktop = $shell.NameSpace(0)
        $items = $desktop.Items()
        
        if ($items.Count -eq 0) {
            Write-Host "No desktop shortcuts found." -ForegroundColor Yellow
            exit 0
        }
        
        Write-Host "Found $($items.Count) desktop items using COM method. Note: This method has limitations." -ForegroundColor Green
        Write-Host "For full functionality, try refreshing desktop (F5) and running script again." -ForegroundColor Cyan
        exit 0
    }
    catch {
        Write-Host "All methods failed. Desktop shortcut rearrangement not possible on this system." -ForegroundColor Red
        exit 1
    }
}

# Get desktop dimensions
$rect = New-Object Win32+RECT
[Win32]::GetWindowRect($desktopListView, [ref]$rect)
$desktopWidth = $rect.Right - $rect.Left
$desktopHeight = $rect.Bottom - $rect.Top

# Get screen dimensions as fallback
if ($desktopWidth -le 0 -or $desktopHeight -le 0) {
    $desktopWidth = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
    $desktopHeight = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height
}

# Get number of desktop items
$itemCount = [Win32]::SendMessage($desktopListView, [Win32]::LVM_GETITEMCOUNT, 0, 0)

if ($itemCount -eq 0) {
    Write-Host "No desktop shortcuts found to rearrange." -ForegroundColor Yellow
    exit 0
}

Write-Host "Found $itemCount desktop shortcuts. Randomizing positions..." -ForegroundColor Green

# Create random number generator
$random = New-Object System.Random

# Define grid parameters
$iconWidth = 80
$iconHeight = 80
$marginX = 10
$marginY = 10

# Calculate grid dimensions
$gridCols = [Math]::Max(1, [Math]::Floor(($desktopWidth - $marginX * 2) / $iconWidth))
$gridRows = [Math]::Max(1, [Math]::Floor(($desktopHeight - $marginY * 2) / $iconHeight))

# Create and shuffle positions
$positions = @()
for ($row = 0; $row -lt $gridRows; $row++) {
    for ($col = 0; $col -lt $gridCols; $col++) {
        $x = $marginX + ($col * $iconWidth)
        $y = $marginY + ($row * $iconHeight)
        $positions += @{ X = $x; Y = $y }
    }
}

# Shuffle positions
for ($i = $positions.Count - 1; $i -gt 0; $i--) {
    $j = $random.Next(0, $i + 1)
    $temp = $positions[$i]
    $positions[$i] = $positions[$j]
    $positions[$j] = $temp
}

# Reposition items
for ($i = 0; $i -lt $itemCount; $i++) {
    if ($i -lt $positions.Count) {
        $pos = $positions[$i]
        $lParam = ($pos.Y -shl 16) -bor ($pos.X -band 0xFFFF)
        [Win32]::SendMessage($desktopListView, [Win32]::LVM_SETITEMPOSITION, $i, $lParam)
    }
}

# Force refresh
[Win32]::SendMessage($desktopListView, [Win32]::WM_COMMAND, 0, 0)

Write-Host "Desktop shortcuts randomized! Press F5 if icons don't refresh immediately." -ForegroundColor Green
