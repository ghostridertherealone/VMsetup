# Desktop Shortcut Randomizer
# Randomly rearranges desktop shortcuts to different grid positions

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
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    
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

function Get-DesktopWindow {
    # Try standard desktop window first
    $progman = [Win32]::FindWindow("Progman", "Program Manager")
    $shellDll = [Win32]::FindWindowEx($progman, [IntPtr]::Zero, "SHELLDLL_DefView", $null)
    $listView = [Win32]::FindWindowEx($shellDll, [IntPtr]::Zero, "SysListView32", "FolderView")
    
    if ($listView -ne [IntPtr]::Zero) {
        return $listView
    }
    
    # Try alternative desktop window (Windows 10 with multiple monitors or special configs)
    $workerW = [IntPtr]::Zero
    do {
        $workerW = [Win32]::FindWindowEx([IntPtr]::Zero, $workerW, "WorkerW", $null)
        if ($workerW -ne [IntPtr]::Zero) {
            $shellDll = [Win32]::FindWindowEx($workerW, [IntPtr]::Zero, "SHELLDLL_DefView", $null)
            if ($shellDll -ne [IntPtr]::Zero) {
                $listView = [Win32]::FindWindowEx($shellDll, [IntPtr]::Zero, "SysListView32", "FolderView")
                if ($listView -ne [IntPtr]::Zero) {
                    return $listView
                }
            }
        }
    } while ($workerW -ne [IntPtr]::Zero)
    
    return [IntPtr]::Zero
}

function Get-ItemCount($hwnd) {
    return [Win32]::SendMessage($hwnd, [Win32]::LVM_GETITEMCOUNT, 0, 0)
}

function Get-ItemPosition($hwnd, $index) {
    $position = [Win32]::SendMessage($hwnd, [Win32]::LVM_GETITEMPOSITION, $index, 0)
    $x = $position -band 0xFFFF
    $y = ($position -shr 16) -band 0xFFFF
    return @{ X = $x; Y = $y }
}

function Set-ItemPosition($hwnd, $index, $x, $y) {
    $lParam = ($y -shl 16) -bor ($x -band 0xFFFF)
    [Win32]::SendMessage($hwnd, [Win32]::LVM_SETITEMPOSITION, $index, $lParam) | Out-Null
}

try {
    Write-Host "Randomizing desktop shortcuts..." -ForegroundColor Green
    
    # Get desktop ListView window
    $desktopHwnd = Get-DesktopWindow
    if ($desktopHwnd -eq [IntPtr]::Zero) {
        throw "Could not find desktop window"
    }
    
    # Get number of items on desktop
    $itemCount = Get-ItemCount $desktopHwnd
    if ($itemCount -eq 0) {
        Write-Host "No shortcuts found on desktop" -ForegroundColor Yellow
        exit
    }
    
    Write-Host "Found $itemCount desktop items"
    
    # Get desktop dimensions
    Add-Type -AssemblyName System.Windows.Forms
    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $desktopWidth = $screen.Width - 100  # Leave margin for icon width
    $desktopHeight = $screen.Height - 100  # Leave margin for icon height
    
    # Generate random positions across entire desktop
    $random = New-Object System.Random
    $positions = @()
    for ($i = 0; $i -lt $itemCount; $i++) {
        $x = $random.Next(50, $desktopWidth)
        $y = $random.Next(50, $desktopHeight)
        $positions += @{ X = $x; Y = $y }
    }
    
    # Apply shuffled positions
    for ($i = 0; $i -lt $itemCount; $i++) {
        Set-ItemPosition $desktopHwnd $i $positions[$i].X $positions[$i].Y
    }
    
    # Refresh desktop
    $progman = [Win32]::FindWindow("Progman", "Program Manager")
    [Win32]::SendMessage($progman, 0x111, 0x7103, 0) | Out-Null
    
    Write-Host "Desktop shortcuts randomized successfully!" -ForegroundColor Green
    
} catch {
    Write-Error "Error: $($_.Exception.Message)"
    exit 1
}
