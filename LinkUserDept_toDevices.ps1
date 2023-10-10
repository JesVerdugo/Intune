install-module Microsoft.Graph.Intune
install-module AzureAD


Connect-MSGraph
Connect-AzureAD

$AzureADUserDepartment = ""
$deviceCategory = ""

Write-Output ""
Write-Output "Script start..."
Write-Output ""


try
{
    $IntuneDevices = Get-IntuneManagedDevice -Filter "operatingSystem eq 'Windows'"
    write-verbose "Success to get Device List"
}
catch
{
    write-Error "Failed to get Device List with error: $_"
}

foreach($IntuneDevice in $IntuneDevices)
{
    $IntuneDeviceID = $IntuneDevice.id

    try
    {

         $primaryUser = (Invoke-MSGraphRequest -HttpMethod GET -Url https://graph.microsoft.com/beta/deviceManagement/managedDevices/$IntuneDeviceID/users).value.userPrincipalName
         Write-Output "Successfully got Primary User `(`"$primaryUser`"`) for device `"$($IntuneDevice.DeviceName)`""
         $AzureADUserDepartment = (Get-AzureADUser -ObjectId $primaryUser).department
         $deviceinformation = Get-IntuneManagedDevice -managedDeviceId $IntuneDeviceID
         $deviceCategory = $($deviceinformation).deviceCategoryDisplayName


         if ($deviceCategory -ne $AzureADUserDepartment)
         {
            $GraphDeviceCategory = Get-IntuneDeviceCategory | Where-Object {$_.DisplayName -like "$AzureADUserDepartment"}
            $GraphDeviceCategoryID = ($GraphDeviceCategory | Select-Object Id).id

            try
            {
                Write-Output "Changing $($IntuneDevice.DeviceName)'s device category from `"$($deviceCategory)`" to `"$($AzureADUserDepartment)`""
                $body = @{ "@odata.id" = "https://graph.microsoft.com/beta/deviceManagement/deviceCategories/$GraphDeviceCategoryID" }
                Invoke-MSGraphRequest -HTTPMethod PUT -Url "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$IntuneDeviceID/deviceCategory/`$ref" -Content $body
                $GraphNewDeviceCategoryDisplayName = (Get-IntuneManagedDevice -ManagedDeviceId $IntuneDeviceID).DeviceCategoryDisplayName
                Write-Output "New Device Category for device `"$($IntuneDevice.DeviceName)`" is `"$($GraphNewDeviceCategoryDisplayName)`""
                Write-Output ""
            }

            catch
            {
                Write-Warning "Failed to update Device Category for device `"$($IntuneDevice.DeviceName)`" with error: $_"

            }


         }
    
        else
        {
            Write-Output "$($IntuneDevice.DeviceName)'s device category `(`"$($deviceCategory)`"`) matches `(`"$primaryUser`"`)'s AAD Department `(`"$($AzureADUserDepartment)`"`)"
            Write-Output ""


        }

         
    }

    catch
    {
       Write-Warning "Failed to get Primary User for device with error: $_"
 
    
    }

    
}



