# Define MAC address (12 hex digits, no colons or dashes)
$NewMAC = "001122AABBCC"  # Customize this

# Get adapter name (assumes it's the one with Internet access)
$adapter = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | Select-Object -First 1

# Check the adapter
Write-Host "Modifying adapter: $($adapter.Name)"

# Set Locally Administered Address (LAA)
Set-NetAdapterAdvancedProperty -Name $adapter.Name -DisplayName "Locally Administered Address" -DisplayValue $NewMAC

# Restart the adapter
Disable-NetAdapter -Name $adapter.Name -Confirm:$false
Start-Sleep -Seconds 2
Enable-NetAdapter -Name $adapter.Name
