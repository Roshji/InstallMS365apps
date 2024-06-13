<#
###################
Authors: M van Rijn - Prodicom 
Remco van Diermen - RvD IT Consulting

# --- UNINSTALL ---
# Some vendors pre-install [Productname], such as Microsoft Surface devices
# Specify all languages to uninstall in file "Uninstall_[Productname]_Languages.xml"
# If no languages need to be removed, only add the required languages to the XML file
# If the uninstall section is not required, remove it

# --- INSTALL ---
# Specify languages in "Install_[Productname].xml" for installation

1. Check if [Productname] is installed
--> YES
Remove languages specified in Uninstall_[Productname]_Languages.xml, end logging and script
--> NO
- Create download folder for Officedeploymenttool
- Download latest version of the Officedeploymenttool
- Unpack Officedeploymenttool
- Download [Productname] via Officedeploymenttool and Install_[Productname].xml to the (package) $Global:DownloadAppFolder
- Install [Productname] via Officedeploymenttool and Install_[Productname].xml

The Officedeploymenttool folder will remain on the device for uninstall
####################
#>

# Vars
$Date = (Get-Date).tostring("yyyy-MM-dd HHmm")
$Global:ProductName = "Office365"
$Global:ExtractAppFolder = "Officedeploymenttool"
$Global:ODTExe = "officedeploymenttool.exe"
$Global:DownloadAppFolder = "C:\ProgramData\Intune\Packages\$($Global:ProductName)"
$LogFile = "C:\Programdata\Microsoft\IntuneManagementExtension\Logs\Logs_$($Global:ProductName)_install_$($Date).log"
$Global:installFolder = "$PSScriptRoot"
$Office365Path = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
$Url = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117"
$Response = Invoke-WebRequest -UseBasicParsing -Uri $url -ErrorAction SilentlyContinue
$Global:SetupEXEFile = "$Global:DownloadAppFolder\$Global:ExtractAppFolder\Setup.exe"
$Global:InstallXML = "Install_$Global:ProductName.xml"
$UnInstallXML = "Uninstall_$($Global:ProductName)_Languages.xml"

# Start Logging
Start-Transcript -Path $LogFile

# Functions
Function Get-OfficeSource {
    Write-host "Setup.exe for uninstall not found. Download ODT package"
    ###########################
    # FOLDER CREATION SECTION #
    ###########################
    # Check if folder exist else create
    If (!(Test-Path $Global:DownloadAppFolder)) {
        # Create directory
        Write-Host "Create ODT package folder: $Global:DownloadAppFolder"
        New-item -path "$Global:DownloadAppFolder" -ItemType "directory"
    }

    ####################
    # DOWNLOAD SECTION #
    ####################
    # Get Download URL of latest Office Deployment Tool (ODT)
    $ODTUri = $response.links | Where-Object {$_.outerHTML -like "*click here to download manually*"}
    $UrlCurrentVerODT = $ODTUri.href
    # Download latest Office Deployment Tool (ODT)
    Write-Host "Downloading latest version of Office Deployment Tool (ODT)."
    Invoke-WebRequest -UseBasicParsing -Uri $UrlCurrentVerODT -OutFile $Global:DownloadAppFolder\$Global:ODTExe
    # Get file version
    Write-Host "Get fileversion Office Deployment Tool (ODT)."
    $Version = (Get-Command $Global:DownloadAppFolder\$Global:ODTExe).FileVersionInfo.FileVersion
    Write-Host "Fileversion Office Deployment Tool (ODT): " $Version
    # Unpack ODT File
    Write-Host "Unpacking file ..."
    $arguments1 = "/quiet /extract:$Global:DownloadAppFolder\$Global:ExtractAppFolder"
    Start-Process -Wait -FilePath "$Global:DownloadAppFolder\$Global:ODTExe" -ArgumentList $arguments1
    Start-sleep -s 5
    # Download
    $arguments1 = "/download $Global:installFolder\$Global:InstallXML" 
    Write-Host "Start download via $Global:SetupEXEFile  $arguments1..."
    Start-Process -Wait -FilePath "$Global:SetupEXEFile" -ArgumentList $arguments1
}

# Check if [Productname] (64 bits) is already installed

If (Test-Path $Office365Path) {
    #####################
    # UNINSTALL SECTION #
    #####################
    Write-Host "$Global:ProductName installation found..."
    # Getting installed components   
    Write-host "Installed $Global:ProductName components..." 
    Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail* | Select-Object DisplayName -expandproperty Displayname

    # Check if original Setup files are present on the system, if not download ODT
    If (!(Test-Path $Global:SetupEXEFile)) {Get-OfficeSource} # Call Function

    # Specify all languages to uninstall in file "Uninstall_[Productname]_Languages.xml"
    # If the languages to uninstall need to be defined first, run the commands above (Get-ItemProperty..) on an identical device and update the file "Uninstall_[Productname]_Languages.xml"
    # Uninstall languages
    Write-Host "Uninstall $Global:ProductName languages...based on XML $Global:installFolder\$UnInstallXML"
    $arguments1 = "/configure $Global:installFolder\$UnInstallXML" 
    Start-Process -Wait -FilePath "$Global:SetupEXEFile" -ArgumentList $arguments1
    
} Else {
    Write-Host "No $Global:ProductName installation found. Download and install $Global:ProductName..."
    Get-OfficeSource # Call Function download ODT

    # Install
    Write-Host "Installing $Global:ProductName..."
    $arguments1 = "/configure $Global:installFolder\$Global:InstallXML" 
    Start-Process -Wait -FilePath "$Global:SetupEXEFile" -ArgumentList $arguments1
}

###################
# CLEANUP SECTION #
###################
Start-sleep -s 5
# Cleanup ODT file. Folder with officedeploymenttool and software will remain for uninstall. Change if needed
Write-Host "Cleanup ODT download file"
Remove-Item "$Global:DownloadAppFolder\$Global:ODTExe"

Start-sleep -s 20

# Stop Logging
Stop-Transcript
