# Define MAC prefixes for Intel and AMD
$IntelPrefixes = @("8086CB", "8086E7", "001B21", "0019D1", "001E37", "002586", "0026B6", "002710", "00270E", "0027EB")
$AMDPrefixes = @("001018", "00E018", "002155", "0050E4", "00D0B7", "001FC6", "0019BB", "001731", "0014A5", "00166F")

# Randomly select between Intel and AMD
$AllPrefixes = $IntelPrefixes + $AMDPrefixes
$RandomPrefix = Get-Random -InputObject $AllPrefixes

# Generate random 6-digit suffix (last 3 bytes)
$RandomSuffix = ""
for ($i = 0; $i -lt 6; $i++) {
    $RandomSuffix += "{0:X}" -f (Get-Random -Minimum 0 -Maximum 16)
}

# Combine prefix and suffix
$NewMAC = $RandomPrefix + $RandomSuffix

Write-Host "Generated MAC: $NewMAC"

# Get adapter name (assumes it's the one with Internet access)
$adapter = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | Select-Object -First 1

# Check the adapter
Write-Host "Modifying adapter: $($adapter.Name)"

# Set Locally Administered Address (LAA)
Set-NetAdapterAdvancedProperty -Name $adapter.Name -DisplayName "Locally Administered Address" -DisplayValue $NewMAC

# Restart the adapter
Disable-NetAdapter -Name $adapter.Name -Confirm:$false
Enable-NetAdapter -Name $adapter.Name

Write-Host "MAC address changed to: $NewMAC"
