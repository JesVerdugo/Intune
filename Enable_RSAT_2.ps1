<# Disable WSUServer value to 1 Run Windows Capability to directly download the components from internet Enable WSUServer value to 0 #>
Write-Host "Adding Components…" -ForegroundColor Green
Get-WindowsCapability -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -Online | Select-Object name | ForEach-Object {Add-WindowsCapability -Name $_.Name -Online}
Set-Content -Path "C:\Windows\debug\Enable_RSAT_Features_Windows10.log" -Value "RSAT Successfully installed"