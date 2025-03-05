<# ###################
Author: M van Rijn - Prodicom

Version 1.1 - 2025-03-03 - Remco van Diermen (RvD IT Consulting)

### CHANGELOG ###
- Updated the Download URL Link
- Added Version checking in section Get-OfficeSource, if an old version of MS365 apps is installed it will be removed first. 

# --- UNINSTALL ---
# Some vendors pre-install [Productname], such as Microsoft Surface devices
# Specify all languages to uninstall in file "Uninstall_[Productname]_Languages.xml"
# If no languages need to be removed, only add the required languages to the XML file
# If the uninstall section is not required, remove it

# --- INSTALL ---
# Specify languages in "Install_[Productname].xml" for installation

1. Check if [Productname] is installed and the version isn't older than x
--> YES
Remove languages specified in Uninstall_[Productname]_Languages.xml, end logging and script
--> NO
- Create download folder for Officedeploymenttool
- Download latest version of the Officedeploymenttool
- Unpack Officedeploymenttool
- Download [Productname] via Officedeploymenttool and Install_[Productname].xml to the (package) $DownloadAppFolder
- Install [Productname] via Officedeploymenttool and Install_[Productname].xml

The Officedeploymenttool folder will remain on the device for uninstall
#################### #>

# Vars
$Date = (Get-Date).tostring("yyyy-MM-dd HHmm")
$ProductName = "Office365"
$ExtractAppFolder = "Officedeploymenttool"
$ODTExe = "officedeploymenttool.exe"
$DownloadAppFolder = "C:\ProgramData\Intune\Packages\$($ProductName)"
$LogFile = "C:\Programdata\Microsoft\IntuneManagementExtension\Logs\Logs_$($ProductName)_install_$($Date).log"
$InstallFolder = "$PSScriptRoot"
$Office365Path = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
$targetVersion = '16.0.18429.20158'
#$Url = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"
# Link changed on Dec 6 2024
$Url = "https://www.microsoft.com/en-us/download/details.aspx?id=49117"
$Response = Invoke-WebRequest -UseBasicParsing -Uri $url -ErrorAction SilentlyContinue
$SetupEXEFile = "$DownloadAppFolder\$ExtractAppFolder\Setup.exe"
$InstallXML = "Install_$ProductName.xml"
$UnInstallXML = "Uninstall_$($ProductName)_Languages.xml"

# Start Logging
Start-Transcript -Path $LogFile

# Functions
Function Get-OfficeSource {
    Write-host "Setup.exe for uninstall not found. Download ODT package"
    ###########################
    # FOLDER CREATION SECTION #
    ###########################
    # Check if folder exist else create
    If (!(Test-Path $DownloadAppFolder)) {
        # Create directory
        Write-Host "Create ODT package folder: $DownloadAppFolder"
        New-item -path "$DownloadAppFolder" -ItemType "directory"
    }

    ####################
    # DOWNLOAD SECTION #
    ####################
    # Get Download URL of latest Office Deployment Tool (ODT)
    #$ODTUri = $response.links | Where-Object {$_.outerHTML -like "*click here to download manually*"}
    #$UrlCurrentVerODT = $ODTUri.href
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
    $arguments1 = "/quiet /extract:$DownloadAppFolder\$ExtractAppFolder"
    Start-Process -Wait -FilePath "$DownloadAppFolder\$ODTExe" -ArgumentList $arguments1
    Start-sleep -s 5
    # Download
    $arguments1 = "/download $installFolder\$InstallXML" 
    Write-Host "Start download via $SetupEXEFile  $arguments1..."
    Start-Process -Wait -FilePath "$SetupEXEFile" -ArgumentList $arguments1
}

# Check if [Productname] (64 bits) is already installed and if version matche

$fileVersion = (Get-Item 'C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeClickToRun.exe').VersionInfo.FileVersion

If ((Test-Path $Office365Path) -and ([version]$fileVersion -ge [version]$targetVersion)) {
    #####################
    # UNINSTALL SECTION #
    #####################
    Write-Host "$ProductName installation found..."
    # Getting installed components   
    Write-host "Installed $ProductName components..." 
    Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail* | Select-Object DisplayName -expandproperty Displayname

    # Check if original Setup files are present on the system, if not download ODT
    If (!(Test-Path $SetupEXEFile)) {Get-OfficeSource} # Call Function

    # Specify all languages to uninstall in file "Uninstall_[Productname]_Languages.xml"
    # If the languages to uninstall need to be defined first, run the commands above (Get-ItemProperty..) on an identical device and update the file "Uninstall_[Productname]_Languages.xml"
    # Uninstall languages
    Write-Host "Uninstall $ProductName languages...based on XML $installFolder\$UnInstallXML"
    $arguments1 = "/configure $installFolder\$UnInstallXML" 
    Start-Process -Wait -FilePath "$SetupEXEFile" -ArgumentList $arguments1
    
} Else {
    Write-Host "No $ProductName installation or older version found. Download and install $ProductName..."
    Write-Host "Currentversion is $fileVersion"
    Get-OfficeSource # Call Function download ODT

    # Install
    Write-Host "Installing $ProductName..."
    $arguments1 = "/configure $installFolder\$InstallXML" 
    Start-Process -Wait -FilePath "$SetupEXEFile" -ArgumentList $arguments1

    Write-Host "Uninstall $ProductName languages...based on XML $installFolder\$UnInstallXML"
    $arguments1 = "/configure $installFolder\$UnInstallXML" 
    Start-Process -Wait -FilePath "$SetupEXEFile" -ArgumentList $arguments1
}

###################
# CLEANUP SECTION #
###################
Start-sleep -s 5
# Cleanup ODT file. Folder with officedeploymenttool and software will remain for uninstall. Change if needed
Write-Host "Cleanup ODT download file"
Remove-Item "$DownloadAppFolder\$ODTExe"

Start-sleep -s 20

# Stop Logging
Stop-Transcript
