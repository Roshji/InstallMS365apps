# Check if App source is downloaded and installed
# An ODT exe file will be present when installed via the new Win32 method
$ProductName = "Office365"
$ODTFolder = "Officedeploymenttool"
$ODTDownloadFolder = "C:\ProgramData\Intune\Packages\$ProductName\$ODTFolder"
$AppFolder = "C:\Program Files\Microsoft Office\root\Office16"
$AppFile = "EXCEL.EXE"

If ((Test-Path $AppFolder\$AppFile) -and (Test-Path $ODTDownloadFolder)) {
    Write-Output "App detected"
    Exit 0
} Else {
    Exit 1
}
