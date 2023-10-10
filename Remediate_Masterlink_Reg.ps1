# Template for Proactive Remediations


#Try-Catch for error handling
Try {
    # After you export the RegKey, be sure you copy/paste it HERE: https://reg2ps.azurewebsites.net/
    # This will create the detection script and the remediation script. 
    
  if((Test-Path -LiteralPath "HKLM:\Software\Microsoft\Office\MS Project\Addins\MasterLink Online") -ne $true) {  New-Item "HKLM:\Software\Microsoft\Office\MS Project\Addins\MasterLink Online" -force -ea SilentlyContinue };
if((Test-Path -LiteralPath "HKLM:\Software\Microsoft\Office\MS Project\Addins\MasterLink Online\MasterLink RegistrationKeys") -ne $true) {  New-Item "HKLM:\Software\Microsoft\Office\MS Project\Addins\MasterLink Online\MasterLink RegistrationKeys" -force -ea SilentlyContinue };
New-ItemProperty -LiteralPath 'HKLM:\Software\Microsoft\Office\MS Project\Addins\MasterLink Online' -Name 'Manifest' -Value 'C:\Program Files (x86)\Matan\MasterLink Client\MasterLink.vsto|vstolocal' -PropertyType String -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKLM:\Software\Microsoft\Office\MS Project\Addins\MasterLink Online' -Name 'Description' -Value 'MasterLink is a MSProject Add-in. ® Matan' -PropertyType String -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKLM:\Software\Microsoft\Office\MS Project\Addins\MasterLink Online' -Name 'FriendlyName' -Value 'MasterLink' -PropertyType String -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKLM:\Software\Microsoft\Office\MS Project\Addins\MasterLink Online' -Name 'LoadBehavior' -Value 3 -PropertyType DWord -Force -ea SilentlyContinue;
New-ItemProperty -LiteralPath 'HKLM:\Software\Microsoft\Office\MS Project\Addins\MasterLink Online\MasterLink RegistrationKeys' -Name 'https://mirati.sharepoint.com/' -Value '6ae8d6ad-1bfb-44f0-8dcc-ac2383389e79' -PropertyType String -Force -ea SilentlyContinue;
    
}Catch{
    #captures and reports the exception errors of the script
    Write-Host $_.Exception
    Exit 2000
}