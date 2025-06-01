$DesktopPath = [Environment]::GetFolderPath("Desktop")
$PublicDesktopPath = [Environment]::GetFolderPath("CommonDesktopDirectory")

# Get all .lnk files from user desktop
Get-ChildItem -Path $DesktopPath -Filter "*.lnk" | 
    Where-Object { $_.Name -ne "Recycle Bin.lnk" } | 
    Remove-Item -Force

# Get all .lnk files from public desktop
Get-ChildItem -Path $PublicDesktopPath -Filter "*.lnk" | 
    Where-Object { $_.Name -ne "Recycle Bin.lnk" } | 
    Remove-Item -Force

Write-Host "Desktop shortcuts deleted (except Recycle Bin)"
