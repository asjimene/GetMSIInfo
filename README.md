# GetMSIInfo
 GetMSIInfo is a simple PowerShell script that adds a right click context menu to grab information for MSI files.

# Installation
 1. Run GetMSIInfo.ps1 with the "-Install" flag to install the appliction to your %LOCALAPPDATA% folder and set the appropriate registry keys, this is a User Based installation.

# Usage
 * Right click an MSI file and click "Get MSI Information" to retrieve the "Property" table from within the MSI.
 * Select one or multiple properties and click OK, the Values for the Properties are copied to your clipboard

# Uninstallation
 1. Run the GetMSIInfo.ps1 script with the "-Uninstall" flag to remove the script and associated registry keys.

# Special Thanks
 Adam Bertram for Get-MsiDatabaseProperties from https://github.com/adbertram/Random-PowerShell-Work/blob/master/Software/Get-MSIDatabaseProperties.ps1 which made pulling the MSI data easy!

 Roger Zander for https://reg2ps.azurewebsites.net/ which made the registry edits (and numerous other things in my life) easy!