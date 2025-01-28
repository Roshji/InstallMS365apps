<# ###################
Author: M van Rijn - Prodicom

Uninstall $Productname based on XML
#################### #>

# Vars
$ProductName = "Office365"
$Date = (Get-Date).tostring("yyyy-MM-dd HHmm")
$LogFile = "C:\Programdata\Microsoft\IntuneManagementExtension\Logs\Logs_$($ProductName)_Uninstall_$($Date).log"
$installFolder = "$PSScriptRoot\"
#$Url = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"
# Link changed on Dec 6 2024
$Url = "https://www.microsoft.com/en-us/download/details.aspx?id=49117"
$Response = Invoke-WebRequest -UseBasicParsing -Uri $url -ErrorAction SilentlyContinue
$DownloadRootFolder = "C:\ProgramData\Intune\Packages\$ProductName"
$ExtractAppFolder = "Officedeploymenttool"
$ODTExe = "officedeploymenttool.exe"
$SetupEXEFile = "$DownloadRootFolder\$DownloadAppFolder\$ExtractAppFolder\Setup.exe"
$UnInstallXML = "Uninstall_$ProductName.xml"

# Start Logging
Start-Transcript -Path $LogFile

# Check if original Setup files are present on the system, if not download ODT
If (!(Test-Path $SetupEXEFile)) {
    Write-host "Setup.exe for uninstall not found. Download ODT package"
    ###########################
    # FOLDER CREATION SECTION #
    ###########################
    # Check if folder exist else create
    If (!(Test-Path $DownloadAppFolder)) {
        # Create directory
        New-item -path "$DownloadAppFolder" -ItemType "directory"
    }

    ####################
    # DOWNLOAD SECTION #
    ####################
    # Get Download URL of latest Office Deployment Tool (ODT)
 #   $ODTUri = $response.links | Where-Object {$_.outerHTML -like "*click here to download manually*"}
 #   $UrlCurrentVerODT = $ODTUri.href
    $ODTUri = $response.Links | Where-Object { $_.href -match "https://download\.microsoft\.com/download/.*/officedeploymenttool_.*\.exe" }
    $UrlCurrentVerODT = $ODTUri.href
 # Download latest Office Deployment Tool (ODT)
    Write-Host "Downloading latest version of Office Deployment Tool (ODT)."
    Invoke-WebRequest -UseBasicParsing -Uri $UrlCurrentVerODT -OutFile $DownloadAppFolder\$ODTExe
    # Get file version
    Write-Host "Get fileversion Office Deployment Tool (ODT)."
    $Version = (Get-Command $DownloadAppFolder\$ODTExe).FileVersionInfo.FileVersion
    Write-Host "Fileversion Office Deployment Tool (ODT): " $Version
    # Unpack ODT File
    Write-Host "Unpacking file ..."
    cd $DownloadAppFolder
    .\$ODTExe /quiet /extract:.\$ExtractAppFolder
    Start-sleep -s 5
}

# Remove
Write-Host "Remove $ProductName..."
$arguments1 = "/Configure $installFolder\$UnInstallXML" 
Start-Process -Wait -FilePath "$SetupEXEFile" -ArgumentList $arguments1

###################
# CLEANUP SECTION #
###################
Start-sleep -s 5
# Cleanup
Write-Host "Cleanup, remove folder $DownloadRootFolder"
Remove-Item -Path "$DownloadRootFolder" -Recurse -Force

# Stop Logging
Stop-Transcript
