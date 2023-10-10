# Created by The Jesus


# Uninstalls Bomgar client older than 2/20/2023 and MSI version
$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

$DetectApp = Get-ChildItem -Path $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "BeyondTrust Remote*" -and $_.InstallDate -lt "20230220" }
$UninstallGUID = $($DetectApp).PSChildName

# In case there are more than one duplicate Bomgar installs we are running this in a loop
foreach ($UninstallGUIDS in $UninstallGUID) {
    $UninstallBomgar = "msiexec.exe /x $UninstallGUIDS /q /log C:\ProgramData\MRTX\Tags\Bomgar\Bomgar_Uninstall.log"
Write-Host "$($UninstallBomgar )"
$UninstallBomgar | cmd
}
# Remove EXE version
$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
$DetectApp = Get-ChildItem -Path $RegPath | Get-ItemProperty | Where-Object { $_.DisplayName -like "Remote Support*"}
if ($DetectApp -ne $null){
Write-Host "Running exe uninstall"
$Filepath =  $($DetectApp).DisplayIcon
                Start-Process -FilePath $Filepath -ArgumentList "-uninstall silent" -WindowStyle Hidden -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 5


}
