<#
.SYNOPSIS
    Script that adds a right click menu to grab info from MSI Files.
.DESCRIPTION
     GetMSIInfo is a simple PowerShell script that adds a right click context menu to grab information for MSI files.
.EXAMPLE
    Install the Script
    GetMSIInfo.ps1 -Install
.EXAMPLE
    Uninstall the Script
    GetMSIInfo.ps1 -Uninstall
.NOTES
    This script is installed in the User Context
    Created by Andrew Jimenez (@asjimene) 2020-04-12
#>

Param (
    # Specifies a path to one or more locations. Unlike the Path parameter, the value of the LiteralPath parameter is
    # used exactly as it is typed. No characters are interpreted as wildcards. If the path includes escape characters,
    # enclose it in single quotation marks. Single quotation marks tell Windows PowerShell not to interpret any
    # characters as escape sequences.
    [Parameter(Mandatory = $false,
        Position = 0,
        ParameterSetName = "LiteralPath",
        ValueFromPipelineByPropertyName = $true,
        HelpMessage = "Literal path to one or more locations.")]
    [Alias("PSPath")]
    [ValidateNotNullOrEmpty()]
    [string[]]
    $LiteralPath,
    # Install Switch Run this Script with the Install Switch to Install the Script and add the Registry changes to your account
    [Parameter(Mandatory = $false)]
    [switch]
    $Install = $false,
    # Uninstall Switch Run this Script with the uninstall Switch to uninstall the Script and remove the Registry changes to your account
    [Parameter(Mandatory = $false)]
    [switch]
    $Uninstall = $false
)

function Get-MsiDatabaseProperties { 
    <# 
    .SYNOPSIS
        This function retrieves properties from a Windows Installer MSI database. 
    .DESCRIPTION
        This function uses the WindowInstaller COM object to pull all values from the Property table from a MSI.
    .EXAMPLE
        Get-MsiDatabaseProperties 'MSI_PATH' 
    .PARAMETER FilePath
        The path to the MSI you'd like to query
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True,
            ValueFromPipeline = $True,
            ValueFromPipelineByPropertyName = $True,
            HelpMessage = 'What is the path of the MSI you would like to query?')]
        [IO.FileInfo[]]$FilePath,
        [Parameter()]
        [String]
        $Table = "Property"
    )

    begin {
        $com_object = New-Object -com WindowsInstaller.Installer
    }

    process {
        try {
            $database = $com_object.GetType().InvokeMember(
                "OpenDatabase",
                "InvokeMethod",
                $Null,
                $com_object,
                @($FilePath.FullName, 0)
            )

            $query = "SELECT * FROM $Table"
            $View = $database.GetType().InvokeMember(
                "OpenView",
                "InvokeMethod",
                $Null,
                $database,
                ($query)
            )

            $View.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $View, $Null)

            $record = $View.GetType().InvokeMember(
                "Fetch",
                "InvokeMethod",
                $Null,
                $View,
                $Null
            )

            $msi_props = @{ }
            while ($record -ne $null) {
                $msi_props[$record.GetType().InvokeMember("StringData", "GetProperty", $Null, $record, 1)] = $record.GetType().InvokeMember("StringData", "GetProperty", $Null, $record, 2)
                $record = $View.GetType().InvokeMember(
                    "Fetch",
                    "InvokeMethod",
                    $Null,
                    $View,
                    $Null
                )
            }

            $msi_props

        }
        catch {
            throw "Failed to get MSI file properties the error was: {0}." -f $_
        }
    }
}


if ($Install) {
    Write-Output "Creating GetMSIInfo folder in LOCALAPPDATA folder"
    New-Item -ItemType Directory -Path $env:LOCALAPPDATA -Name "GetMSIInfo" -ErrorAction SilentlyContinue

    Write-Output "Copying Script to GetMSIInfo Folder"
    Copy-Item $PSScriptRoot\GetMSIInfo.ps1 -Destination "$env:LOCALAPPDATA\GetMSIInfo\GetMSIInfo.ps1" -ErrorAction SilentlyContinue
    
    # Reg2CI (c) 2020 by Roger Zander
    if ((Test-Path -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\.msi") -ne $true) { New-Item "HKCU:\Software\Classes\SystemFileAssociations\.msi" -force -ea SilentlyContinue };
    if ((Test-Path -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell") -ne $true) { New-Item "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell" -force -ea SilentlyContinue };
    if ((Test-Path -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell\Get MSI Information") -ne $true) { New-Item "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell\Get MSI Information" -force -ea SilentlyContinue };
    if ((Test-Path -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell\Get MSI Information\command") -ne $true) { New-Item "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell\Get MSI Information\command" -force -ea SilentlyContinue };
    New-ItemProperty -LiteralPath 'HKCU:\Software\Classes\SystemFileAssociations\.msi\shell\Get MSI Information' -Name '(default)' -Value "Get MSI Information" -PropertyType String -Force -ea SilentlyContinue;
    New-ItemProperty -LiteralPath 'HKCU:\Software\Classes\SystemFileAssociations\.msi\shell\Get MSI Information\command' -Name '(default)' -Value "C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command `"$env:LOCALAPPDATA\GetMSIInfo\GetMSIInfo.ps1`" -LiteralPath '%1'" -PropertyType String -Force -ea SilentlyContinue;

    Write-Output "Installation Complete"
    Pause
}

if ($Uninstall) {
    Write-Output "Removing Script from LOCALAPPDATA"
    Remove-item "$env:LOCALAPPDATA\GetMSIInfo" -Force -Recurse -ErrorAction SilentlyContinue

    Write-Output "Cleaning Up Registry"
    if ((Test-Path -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell\Get MSI Information") -eq $true) { 
        Remove-Item "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell\Get MSI Information" -force -Recurse -ea SilentlyContinue 
    }
    if ([System.String]::IsNullOrEmpty((Get-ChildItem "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell"))) {
        Remove-Item "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell" -force -Recurse -ea SilentlyContinue 
    }
    if ([System.String]::IsNullOrEmpty((Get-ChildItem "HKCU:\Software\Classes\SystemFileAssociations\.msi"))) {
        Remove-Item "HKCU:\Software\Classes\SystemFileAssociations\.msi" -force -Recurse -ea SilentlyContinue 
    }

    Write-Output "Uninstallation Complete!"
    Pause
}

if ((-not $Install) -and (-not $Uninstall)) {
    $SelectedInfo = Get-MsiDatabaseProperties -FilePath $LiteralPath | Out-GridView -Title "MSI Database Properties for $LiteralPath" -OutputMode Multiple
    $SelectedInfo.Value | Set-Clipboard
}