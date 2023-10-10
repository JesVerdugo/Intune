##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uO1
##Nc3NCtDXTlaDjqbQ7ghi60LgUXwqYsnWsLWoysys8e2h62vQSpV0
##Kd3HFJGZHWLWoLaVvnQnhQ==
##LM/RF4eFHHGZ7/K1
##K8rLFtDXTiW5
##OsHQCZGeTiiZ4dI=
##OcrLFtDXTiW5
##LM/BD5WYTiiZ4tI=
##McvWDJ+OTiiZ4tI=
##OMvOC56PFnzN8u+VslQ=
##M9jHFoeYB2Hc8u+VslQ=
##PdrWFpmIG2HcofKIo2QX
##OMfRFJyLFzWE8uK1
##KsfMAp/KUzWI0g==
##OsfOAYaPHGbQvbyVvnQmqxigEiZ7Dg==
##LNzNAIWJGmPcoKHc7Do3uAu/DDtlPovL2Q==
##LNzNAIWJGnvYv7eVvnRF4EbhVG1rXdGaq6KuyoaM8OPir0U=
##M9zLA5mED3nfu77Q7TV64AuzAgg=
##NcDWAYKED3nfu77Q7TV64AuzAgg=
##OMvRB4KDHmHQvbyVvnQX
##P8HPFJGEFzWE8tI=
##KNzDAJWHD2fS8u+Vgw==
##P8HSHYKDCX3N8u+VyTVk6kWubG8+d8CV2Q==
##LNzLEpGeC3fMu77Ro2k3hQ==
##L97HB5mLAnfMu77Ro2k3hQ==
##P8HPCZWEGmaZ7/K1
##L8/UAdDXTlaDjqbQ7ghC60LgUXwqYsmkiqKm1pW18e3TiyrQR45aTExy9g==
##Kc/BRM3KXhU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba
# Script to remove StartingPoint
#Author Jason Nguyen
#version 1.3.2 modified 03/09/2023 @ 3:51 PM
#Minor update to correcting bugs and rewreote logging output for easier trackign and readibility

$ErrorActionPreference = 'silentlycontinue'

# Function to test if registry value exists
function Test-RegistryValue 
{
    param (
        [parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Path,
        [parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Value,
        [switch]$ShowValue
    )
    try {
        $Values = Get-ItemProperty -Path $Path | select-object -ExpandProperty $Value -ErrorAction Stop 
        if ($ShowValue) {
            $Values
        }
        else {
            $true
        }
         
    }
    catch {
        $false
    }
}


# Add a pause for testing - Usage: pause "Press any key to continue"
Function pause ($message)
{
    # Check if running Powershell ISE
    if ($psISE)
    {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("$message")
    }
    else
    {
        Write-Host "$message" -ForegroundColor Yellow
        $x = $host.ui.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}


#Creating logoutput and filenames, catch requires an error output signal 
Function Write-Log
{
	param (
        [Parameter(Mandatory=$True)]
        [array]$LogOutput,
        [Parameter(Mandatory=$True)]
        [string]$Path
	)
	$currentDate = (Get-Date -UFormat "%d-%m-%Y")
	$currentTime = (Get-Date -UFormat "%T")
	$logOutput = $logOutput -join (" ")
	"[$currentDate $currentTime] $logOutput" | Out-File $Path -Append

    <# ---
        An example on how to use it:

        try 
            {
                something right!
                copy-item ... -ErrorAction 'Stop'
            }
            Catch
                {
                    write-log -LogOutput ("Failed to do something right:  {0}" -f $_) -Path $LogFile
                }
    -- #>
}



#Set path and variables
$locationToCheck = "C:\StartingPoint"
$fileToCheck = "StartingPoint.zip" 
$fulllocationfileToCheck = $locationToCheck + "\" + $fileToCheck
$createAppdataFolder = "C:\Appdata"
$completeFlagFolder = "C:\Appdata\StartingPoint_Flag"
$combinedfilefolder = $completeFlagFolder+"\"+$addVersionFlagFile
$checkFileSuccess = "Author.dotm"
$authorPath = $locationToCheck+"\"+$checkFileSuccess
$addVersionFlagFile = "v.txt"
$currentTime = Get-Date -format "MMM-dd-yyyy HH:mm:ss"
$flagmessage = "$currentTime Deployed StartingPoint"
$userAppData = [System.Environment]::GetEnvironmentVariable('appdata')
$userTemp = [System.Environment]::GetEnvironmentVariable('temp')
$templateStartup = "\Microsoft\Templates"
$startuppath = $userAppData+$templateStartup
$LogFolder = $userTemp
$LogFile = $completeFlagFolder + "\" + (Get-Date -UFormat "%d-%m-%Y") + "-DeployStartingPoint.log"
$logFileCleanup = $LogFolder + "\" +  "CleanupStartingPoint.log"
$scriptWorkingPath = (Get-Location).Path
$scriptCheckZipFile = $scriptWorkingPath  + $fileToCheck
$scriptLocation = $PSScriptRoot
$hostname = $env:computername  
$startingpointURL = "https://endpointfile.blob.core.windows.net/startingpoint/StartingPoint.zip?sv=2021-10-04&st=2023-02-01T23%3A14%3A18Z&se=2024-02-02T23%3A14%3A00Z&sr=b&sp=r&sig=lObGUhfzI7CZlKyLKpbQ8FtgfVtMdDQAF3oa1CFrIqE%3D"
$wordStartup = "$env:USERPROFILE\AppData\Roaming\Microsoft\Word\STARTUP"
$wordUserTemplates = "$env:USERPROFILE\AppData\Roaming\Microsoft\Templates"
$userDesktop = "$HOME\OneDrive - Mirati Therapeutics\Desktop"
$searchFoldersStart = $wordStartup
$searchFoldersTemplates = $wordUserTemplates
$searchFoldersSDesktop = $userDesktop

#######################################################################################################################
######------------------- START CLEAN UP ---------------------------------------------------------------###############
#######################################################################################################################

#First we clean up old settings and previous deployment
Write-Host "--==Start Cleanup and searching for known folders for StartingPoint related files...==--`n" -ForegroundColor Green
Write-Log -LogOutput ("Searching known folders for StartingPoint related files...`n") -Path $logFileCleanup

# Search for files in the specified folder and delete the files
Write-Host "1. Searching Word Startup folder.`n" -ForegroundColor Cyan
Write-Log -LogOutput ("1. Searching Word Startup folder.`n") -Path $logFileCleanup
Sleep 1

$workFiles = Get-ChildItem $searchFoldersStart -Recurse -Include "author*.dotm","author*.lnk","spfaq*.pdf","spfaq*.lnk","StartingPoint Style Definition Guide*.pdf","StartingPoint Style Definition Guide*.lnk","StartingPoint User Manual*.pdf","StartingPoint User Manual*.lnk","Tipsheet*.pdf","Tipsheet*.lnk","StartingPointSettings*.ini","StartingPointSettings*.lnk","StartingPoint*.pdf","StartingPoint*.lnk" | Where-Object { $_.PSIsContainer -eq $false }
Sleep 1

    If (!($workFiles -eq $null))
    {
        foreach ($fileFound in $workFiles)
        {
            Remove-Item $fileFound -Force -Verbose -ErrorAction SilentlyContinue
            Write-Host "Found related file(s). File being removed:`n $fileFound`n" -ForegroundColor Yellow
            Write-Log -LogOutput ("Found related file(s). File being removed:`n $fileFound`n") -Path $logFileCleanup  
        }   
    }
    else
        {
            Write-Host "No files found that is related to Author to be removed.`n" -ForegroundColor Yellow
            Write-Log -LogOutput ("No files found that is related to Author to be removed.`n") -Path $logFileCleanup
        }
Sleep 1

Clear-Variable -Name workFiles

Write-Host "2. Searching Word User Template folder.`n" -ForegroundColor Cyan
Write-Log -LogOutput ("2. Searching Word User Template folder.`n") -Path $logFileCleanup
Sleep 1

$workFiles = Get-ChildItem $searchFoldersTemplates -Recurse -Include "Normal.dotm, author*.dotm","author*.lnk","spfaq*.pdf","spfaq*.lnk","StartingPoint Style Definition Guide*.pdf","StartingPoint Style Definition Guide*.lnk","StartingPoint User Manual*.pdf","StartingPoint User Manual*.lnk","Tipsheet*.pdf","Tipsheet*.lnk","StartingPointSettings*.ini","StartingPointSettings*.lnk","StartingPoint*.pdf","StartingPoint*.lnk" | Where-Object { $_.PSIsContainer -eq $false }
Sleep 1

    If (!($workFiles -eq $null))
    {
        foreach ($fileFound in $workFiles)
        {
            Remove-Item $fileFound -Force -Confirm $false -Verbose -ErrorAction SilentlyContinue
            Write-Host "Found related file(s). File being removed:`n $fileFound`n" -ForegroundColor Yellow
            Write-Log -LogOutput ("Found related file(s). File being removed:`n $fileFound`n") -Path $logFileCleanup  
        }   
    }
    else
        {
            Write-Host "No files found that is related to Author to be removed.`n" -ForegroundColor Yellow
            Write-Log -LogOutput ("No files found that is related to Author to be removed.`n") -Path $logFileCleanup
        }
Sleep 1

Clear-Variable -Name workFiles
Write-Host "3. Searching User Desktop folder.`n" -ForegroundColor Cyan
Write-Log -LogOutput ("3. Searching User Desktop folder`n") -Path $logFileCleanup
Sleep 1

$workFiles = Get-ChildItem $searchFoldersSDesktop -Recurse -Include "author*.dotm","author*.lnk","spfaq*.pdf","spfaq*.lnk","StartingPoint Style Definition Guide*.pdf","StartingPoint Style Definition Guide*.lnk","StartingPoint User Manual*.pdf","StartingPoint User Manual*.lnk","Tipsheet*.pdf","Tipsheet*.lnk","StartingPointSettings*.ini","StartingPointSettings*.lnk","StartingPoint*.pdf","StartingPoint*.lnk" | Where-Object { $_.PSIsContainer -eq $false }
Sleep 1

    If (!($workFiles -eq $null))
    {
        foreach ($fileFound in $workFiles)
        {
            Remove-Item $fileFound -Force -Verbose -ErrorAction SilentlyContinue
            Write-Host "Found related file(s). File being removed:`n $fileFound`n" -ForegroundColor Yellow
            Write-Log -LogOutput ("Found related file(s). File being removed:`n $fileFound`n") -Path $logFileCleanup  
        }   
    }
    else
        {
            Write-Host "No files found that is related to Author to be removed.`n" -ForegroundColor Yellow
            Write-Log -LogOutput ("No files found that is related to Author to be removed.`n") -Path $logFileCleanup
        }
Sleep 1



# Set starting location into Registry
$CWD = [Environment]::CurrentDirectory
Set-Location "HKCU:\"


Write-Host "--==Searching the registry for known StartingPoint Values...==--`n" -ForegroundColor Cyan
Write-Log -LogOutput ("Searching the registry for known StartingPoint Values...`n") -Path $logFileCleanup

#Searches Trusted Locations for Regulatory Resources Property value and deleted the key
Write-Host "4. Removing Regulatory Resources - Author Setup Locations from registry`n" -ForegroundColor Yellow
Write-Log -LogOutput ("4. Removing Regulatory Resources - Author Setup Locations from registry`n") -Path $logFileCleanup

Get-ChildItem -Path 'HKCU:\Software\Microsoft\Office\16.0\Word\Security\Trusted Locations' -Recurse | 
ForEach-Object {
    $key = Get-Item $_.PsPath
    if ($key.GetValueNames() | ForEach-Object { $key.GetValue($_) } | Where-Object { $_ -like "*Regulatory Resources - Author Setup*" }) {
        Remove-Item $key.PSPath -Recurse -Force -Verbose -ErrorAction SilentlyContinue
        Write-Host "Found Location key for Regulatory Resources - Author Setup and removing:`n $key`n" -ForegroundColor Yellow
        Write-Log -LogOutput ("Found Location key for Regulatory Resources - Author Setup and removing:`n $key`n") -Path $logFileCleanup
    }
}
Sleep 1

# Will remove previous install per Jason's deployment script
#Set remove-itemproperty path so that it doesn't include the running drive path eg C:\

# Remove VBATrustModify 
Write-Host "5. Remove Digitally Signed Macros from registry`n" -ForegroundColor Yellow
Remove-ItemProperty -Path "SOFTWARE\Microsoft\Office\16.0\Word\Security" -Name "VBAWarnings" -Force -Verbose -ErrorAction SilentlyContinue
Write-Log -LogOutput ("5. Remove Digitally Signed Macros from registry`n") -Path $logFileCleanup
Sleep 1

# Remove SharedTemplateModify
Write-Host "6. Removing Shared Workgroup - StartingPoint from registry`n" -ForegroundColor Yellow
Remove-ItemProperty -Path "SOFTWARE\Microsoft\Office\16.0\Common\General" -Name "SharedTemplates" -Force -Verbose -Verbose -ErrorAction SilentlyContinue
Write-Log -LogOutput ("6. Removing c:\StartingPoint as Shared Workgroup from registry`n") -Path $logFileCleanup
Sleep 1


#Searches Trusted Locations for StartingPoint Property value and deleted the key
Write-Host "7. Removing Trusted StartingPoint Locations from registry`n" -ForegroundColor Yellow
Write-Log -LogOutput ("7. Removing Trusted StartingPoint Locations from registry`n") -Path $logFileCleanup 
Get-ChildItem -Path 'HKCU:\Software\Microsoft\Office\16.0\Word\Security\Trusted Locations' -Recurse | 
ForEach-Object {
    $key = Get-Item $_.PsPath
    if ($key.GetValueNames() | ForEach-Object { $key.GetValue($_) } | Where-Object { $_ -like "*StartingPoint*" }) {
        Remove-Item $key.PSPath -Recurse -Force -Verbose -ErrorAction SilentlyContinue
        Write-Host "Found Location key for StartingPoint and removing:`n $key`n" -ForegroundColor Yellow
        Write-Log -LogOutput ("Found Location key for StartingPoint and removing:`n $key`n") -Path $logFileCleanup
    }
}
Sleep 1

# Remove developer tool
Write-Host "8. Removing Developer Tools to the Ribbon`n" -ForegroundColor Yellow
Remove-ItemProperty -Path "Software\Microsoft\Office\16.0\Word\Options" -Name "DeveloperTools" -Force -Verbose -Verbose -ErrorAction SilentlyContinue
Write-Log -LogOutput ("8. Removing Developer Tools to the Ribbon`n") -Path $logFileCleanup
Sleep 1

#Create Shortcut to autoload Author.dotm
Write-Host "9. Delete Template Autoload shortcut for Author.dotm in Word User Templates`n" -ForegroundColor Yellow
Remove-Item -Path "$env:USERPROFILE\AppData\Roaming\Microsoft\Templates\Author shortcut.lnk" -Verbose -ErrorAction SilentlyContinue
Write-Log -LogOutput ("9. Delete Template Autoload shortcut for Author.dotm in Word User Templates`n") -Path $logFileCleanup
Sleep 1

#Delete FlagFolder
#Set location back
Set-Location $CWD

Write-Host "10. Delete Flag Folder`n" -ForegroundColor Yellow
Remove-Item -LiteralPath "C:\Appdata" -Force -Recurse -Verbose -ErrorAction SilentlyContinue
Write-Log -LogOutput ("10. Delete Flag Folder`n") -Path $logFileCleanup
Sleep 1

Write-Host "11. Delete StartingPoint Folder`n" -ForegroundColor Yellow
Remove-Item -LiteralPath "C:\StartingPoint" -Force -Recurse -Verbose -ErrorAction SilentlyContinue
Write-Log -LogOutput ("11. Delete StartingPoint Folder`n") -Path $logFileCleanup
Sleep 1

Write-Host "12. Done removing deployment settings`n" -ForegroundColor Green
Write-Log -LogOutput ("12. Done removing deployment settings`n") -Path $logFileCleanup
Sleep 1

Write-Host "13. You can review cleanup log here: $logFileCleanup  `n" -ForegroundColor Cyan
Sleep 5

Write-Log -LogOutput ("Exit 0") -Path $logFileCleanup
Exit 0