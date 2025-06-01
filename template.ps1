# Taskbar Randomizer - Randomly rearranges pinned taskbar shortcuts

try {
    # Registry path for taskbar layout
    $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband"
    
    # Get current taskbar favorites
    $favoritesKey = Get-ItemProperty -Path $regPath -Name "Favorites" -ErrorAction SilentlyContinue
    
    if ($favoritesKey -and $favoritesKey.Favorites) {
        # Convert binary data to byte array
        $favoritesData = $favoritesKey.Favorites
        
        # Each pinned item has a specific structure in the binary data
        $itemSize = 32  # Approximate size of each taskbar item entry
        $numItems = [Math]::Floor($favoritesData.Length / $itemSize)
        
        if ($numItems -gt 1) {
            # Create array of item indices
            $indices = 0..($numItems - 1)
            
            # Shuffle the indices using Fisher-Yates algorithm
            for ($i = $indices.Length - 1; $i -gt 0; $i--) {
                $j = Get-Random -Maximum ($i + 1)
                $temp = $indices[$i]
                $indices[$i] = $indices[$j]
                $indices[$j] = $temp
            }
            
            # Create new shuffled data array
            $newData = New-Object byte[] $favoritesData.Length
            
            # Copy shuffled items to new array
            for ($i = 0; $i -lt $numItems; $i++) {
                $sourceIndex = $indices[$i] * $itemSize
                $destIndex = $i * $itemSize
                $copyLength = [Math]::Min($itemSize, $favoritesData.Length - $sourceIndex)
                
                [Array]::Copy($favoritesData, $sourceIndex, $newData, $destIndex, $copyLength)
            }
            
            # Write the shuffled data back to registry
            Set-ItemProperty -Path $regPath -Name "Favorites" -Value $newData -Type Binary
            
            # Restart Explorer to apply changes
            Stop-Process -Name "explorer" -Force
            Start-Sleep -Seconds 2
            Start-Process "explorer.exe"
        }
    }
}
catch {
    Write-Error "Failed to randomize taskbar: $($_.Exception.Message)"
}

[System.GC]::Collect()
