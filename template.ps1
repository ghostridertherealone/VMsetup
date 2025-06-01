# Taskbar Randomizer - Randomly rearranges pinned taskbar shortcuts

try {
    # Use COM object to manipulate taskbar
    $shell = New-Object -ComObject Shell.Application
    $folder = $shell.NameSpace("shell:::{4234d49b-0245-4df3-b780-3893943456e1}")
    
    if ($folder) {
        $items = @()
        for ($i = 0; $i -lt $folder.Items().Count; $i++) {
            $items += $folder.Items().Item($i)
        }
        
        if ($items.Count -gt 1) {
            # Shuffle items using Fisher-Yates algorithm
            for ($i = $items.Count - 1; $i -gt 0; $i--) {
                $j = Get-Random -Maximum ($i + 1)
                $temp = $items[$i]
                $items[$i] = $items[$j]
                $items[$j] = $temp
            }
            
            # Unpin all items
            foreach ($item in $items) {
                $verb = $item.Verbs() | Where-Object { $_.Name -match "Unpin|Remove" }
                if ($verb) { $verb.DoIt() }
            }
            
            Start-Sleep -Seconds 1
            
            # Re-pin in shuffled order
            foreach ($item in $items) {
                $verb = $item.Verbs() | Where-Object { $_.Name -match "Pin" }
                if ($verb) { $verb.DoIt() }
            }
        }
    }
}
catch {
    Write-Error "Failed to randomize taskbar: $($_.Exception.Message)"
}

# Clean up COM objects
if ($shell) {
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
}

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
