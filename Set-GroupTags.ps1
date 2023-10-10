Connect-MgGraph
$clientId = "72209ec4-ab54-4de9-a95b-9635c8285aa2" #Provide the Client ID
$clientSecret = "WKH8Q~YsZAnIlxrUzO3RIYQocQYR1T.7zyC.-cFP" # Provide the ClientSecret
$ourTenantId = "f535743e-b6ee-4533-8cf5-b1e76dd9ae2b" #Specify the TenatID

$Resource = "deviceManagement/windowsAutopilotDeviceIdentities"
$Resource = "deviceManagement/managedDevices"
$graphApiVersion = "Beta"
$uri = "https://graph.microsoft.com/$graphApiVersion/$($resource)"
$authority = "https://login.microsoftonline.com/$ourTenantId"
Update-MSGraphEnvironment -AppId $clientId -Quiet
Update-MSGraphEnvironment -AuthUrl $authority -Quiet
Connect-MSGraph -ClientSecret $clientSecret

$Grouptag = "EUS" #Specify the GroupTag Here
$SerialNumbers = Get-Content -Path "C:\scripts\SerialNumber.txt" #Provide the list of devices you want to check the GroupTag
foreach ($Serial in $SerialNumbers)
{
Get-AutopilotDevice -serial $Serial | Set-AutopilotDevice -groupTag $Grouptag
}
