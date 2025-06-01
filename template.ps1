# Desktop Shortcut Randomizer - Direct Approach
# Moves desktop shortcuts to random positions

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Text;

public class Desktop {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
    
    [DllImport("user32.dll")]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    
    [DllImport("user32.dll")]
    public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
    
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    
    [DllImport("user32.dll")]
    public static extern bool InvalidateRect(IntPtr hWnd, IntPtr lpRect, bool bErase);
    
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left, Top, Right, Bottom;
    }
    
    [StructLayout(LayoutKind.Sequential)]
    public struct POINT {
        public int X, Y;
    }
    
    public const uint LVM_GETITEMCOUNT = 0x1004;
    public const uint LVM_GETITEMPOSITION = 0x1010;
    public const uint LVM_SETITEMPOSITION = 0x100F;
    public const uint LVM_ARRANGE = 0x1016;
    public const uint LVA_DEFAULT = 0x0000;
    public const uint SWP_NOSIZE = 0x0001;
    public const uint SWP_NOZORDER = 0x0004;
}
"@

# Get desktop ListView handle
function Get-DesktopHandle {
    # Try multiple methods for Windows 10/11
    $handles = @()
    
    # Method 1: Classic Progman
    $progman = [Desktop]::FindWindow("Progman", "Program Manager")
    if ($progman -ne [IntPtr]::Zero) {
        $defview = [Desktop]::FindWindowEx($progman, [IntPtr]::Zero, "SHELLDLL_DefView", $null)
        if ($defview -ne [IntPtr]::Zero) {
            $listview = [Desktop]::FindWindowEx($defview, [IntPtr]::Zero, "SysListView32", $null)
            if ($listview -ne [IntPtr]::Zero) { $handles += $listview }
        }
    }
    
    # Method 2: WorkerW windows
    $workerw = [IntPtr]::Zero
    do {
        $workerw = [Desktop]::FindWindowEx([IntPtr]::Zero, $workerw, "WorkerW", $null)
        if ($workerw -ne [IntPtr]::Zero) {
            $defview = [Desktop]::FindWindowEx($workerw, [IntPtr]::Zero, "SHELLDLL_DefView", $null)
            if ($defview -ne [IntPtr]::Zero) {
                $listview = [Desktop]::FindWindowEx($defview, [IntPtr]::Zero, "SysListView32", $null)
                if ($listview -ne [IntPtr]::Zero) { $handles += $listview }
            }
        }
    } while ($workerw -ne [IntPtr]::Zero)
    
    # Return handle with most items
    $bestHandle = [IntPtr]::Zero
    $maxItems = 0
    
    foreach ($handle in $handles) {
        $itemCount = [Desktop]::SendMessage($handle, [Desktop]::LVM_GETITEMCOUNT, [IntPtr]::Zero, [IntPtr]::Zero)
        if ($itemCount -gt $maxItems) {
            $maxItems = $itemCount
            $bestHandle = $handle
        }
    }
    
    return @{ Handle = $bestHandle; Count = $maxItems }
}

# Get desktop info
$desktopInfo = Get-DesktopHandle
$desktopHandle = $desktopInfo.Handle
$itemCount = $desktopInfo.Count

if ($desktopHandle -eq [IntPtr]::Zero -or $itemCount -eq 0) {
    Write-Host "No desktop shortcuts found or unable to access desktop." -ForegroundColor Red
    exit 1
}

Write-Host "Found $itemCount desktop items. Randomizing..." -ForegroundColor Green

# Get screen dimensions
$screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$screenWidth = $screen.Width
$screenHeight = $screen.Height

# Generate random positions
$random = New-Object Random
$iconSize = 75
$margin = 50

# Create grid of possible positions
$cols = [Math]::Floor(($screenWidth - $margin * 2) / $iconSize)
$rows = [Math]::Floor(($screenHeight - $margin * 2) / $iconSize)

$positions = @()
for ($r = 0; $r -lt $rows; $r++) {
    for ($c = 0; $c -lt $cols; $c++) {
        $x = $margin + ($c * $iconSize)
        $y = $margin + ($r * $iconSize)
        $positions += @{X = $x; Y = $y}
    }
}

# Shuffle positions
for ($i = $positions.Count - 1; $i -gt 0; $i--) {
    $j = $random.Next($i + 1)
    $temp = $positions[$i]
    $positions[$i] = $positions[$j]
    $positions[$j] = $temp
}

# Apply new positions
for ($i = 0; $i -lt $itemCount -and $i -lt $positions.Count; $i++) {
    $pos = $positions[$i]
    $x = $pos.X
    $y = $pos.Y
    
    # Pack coordinates into lParam
    $lParam = [IntPtr]($y -shl 16 -bor ($x -band 0xFFFF))
    
    # Set item position
    $result = [Desktop]::SendMessage($desktopHandle, [Desktop]::LVM_SETITEMPOSITION, [IntPtr]$i, $lParam)
}

# Force desktop refresh
[Desktop]::InvalidateRect($desktopHandle, [IntPtr]::Zero, $true)
[Desktop]::SendMessage($desktopHandle, [Desktop]::LVM_ARRANGE, [IntPtr][Desktop]::LVA_DEFAULT, [IntPtr]::Zero)

Write-Host "Desktop shortcuts randomized! Press F5 to refresh if needed." -ForegroundColor Green

# Optional: Restart explorer if changes don't appear
$choice = Read-Host "Restart Windows Explorer to ensure changes? (y/n)"
if ($choice -eq 'y' -or $choice -eq 'Y') {
    Stop-Process -Name explorer -Force
    Start-Process explorer
}
