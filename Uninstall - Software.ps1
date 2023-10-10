$softwareDisplayName = "MasterLink Client"

# Get the product code for the specified software display name
$software = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -eq $softwareDisplayName }

if ($software) {
    $productCode = $software.IdentifyingNumber

    # Uninstall the software using the product code
    $uninstallArgs = "/x $productCode /qn"
    $uninstallProcess = Start-Process -FilePath "msiexec.exe" -ArgumentList $uninstallArgs -Wait -PassThru

    if ($uninstallProcess.ExitCode -eq 0) {
        Write-Host "Software uninstalled successfully."
    } else {
        Write-Host "Failed to uninstall the software. Exit code: $($uninstallProcess.ExitCode)"
    }
} else {
    Write-Host "Software not found."
}
