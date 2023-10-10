#Try-Catch for error handling
Try {
    # After you export the RegKey, be sure you copy/paste it HERE: https://reg2ps.azurewebsites.net/
    # This will create the detection script and the remediation script. 

	if(-NOT (Test-Path -LiteralPath "HKLM:\Software\Microsoft\Office\MS Project\Addins\MasterLink Online")){ Exit 1 };
	if(-NOT (Test-Path -LiteralPath "HKLM:\Software\Microsoft\Office\MS Project\Addins\MasterLink Online\MasterLink RegistrationKeys")){ Exit 1 };
	if((Get-ItemPropertyValue -LiteralPath 'HKLM:\Software\Microsoft\Office\MS Project\Addins\MasterLink Online' -Name 'Manifest' -ea SilentlyContinue) -eq 'C:\Program Files (x86)\Matan\MasterLink Client\MasterLink.vsto|vstolocal') {  } else { Exit 1 };
	if((Get-ItemPropertyValue -LiteralPath 'HKLM:\Software\Microsoft\Office\MS Project\Addins\MasterLink Online' -Name 'Description' -ea SilentlyContinue) -eq 'MasterLink is a MSProject Add-in. ® Matan') {  } else { Exit 1 };
	if((Get-ItemPropertyValue -LiteralPath 'HKLM:\Software\Microsoft\Office\MS Project\Addins\MasterLink Online' -Name 'FriendlyName' -ea SilentlyContinue) -eq 'MasterLink') {  } else { Exit 1 };
	if((Get-ItemPropertyValue -LiteralPath 'HKLM:\Software\Microsoft\Office\MS Project\Addins\MasterLink Online' -Name 'LoadBehavior' -ea SilentlyContinue) -eq 3) {  } else { Exit 1 };
	if((Get-ItemPropertyValue -LiteralPath 'HKLM:\Software\Microsoft\Office\MS Project\Addins\MasterLink Online\MasterLink RegistrationKeys' -Name 'https://mirati.sharepoint.com/' -ea SilentlyContinue) -eq '6ae8d6ad-1bfb-44f0-8dcc-ac2383389e79') {  } else { Exit 1 };
    
}Catch{
    #captures and reports the exception errors of the script
    Write-Host $_.Exception
    Exit 2000
}