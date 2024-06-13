## Powershell Script for checking your MS365 Tenant for licensing availability

Win32 App

Name: Microsoft Office 365
Description: Microsoft Office 365
Publisher: Microsoft
App Version: 365
Allow available uninstall: No
Category: Business
Show this as a featured app in the Company Portal: No
Install behavior: System

Install command: powershell.exe -ExecutionPolicy Bypass -File Install_Office365.ps1

Uninstall command: powershell.exe -ExecutionPolicy Bypass -File Uninstall_Office365.ps1

Detection rule: [Script]
Detection_Office365.ps1
Run script as 32-bit process on 64-bit clients: No
Enforce script signature check and run script silently: No


