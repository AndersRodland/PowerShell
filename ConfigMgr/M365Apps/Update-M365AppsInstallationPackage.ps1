<#
    .SYNOPSIS
    Update M365 Apps / Office 365 installation package to latest version.
    .DESCRIPTION
    Update M365 Apps / Office 365 installation package to latest version.

    .PARAMETER Siteserver
    Define the hostname of the Microsoft Endpoint Manager Configuration Manager site server.

    .PARAMETER Sitecode
    Define the sitecode of the Microsoft Endpoint Manager Configuration Manager site.

    .PARAMETER ApplicationName
    Define the application name of the Microsoft 365 Apps application in Microsoft Endpoint Manager Configuration Manager.

    .PARAMETER Path
    Define the path of the Microsoft 365 Apps application content source

    .EXAMPLE
    .\Update-M365AppsInstallationPackage.ps1 -Siteserver "MEMCM01.andersrodland.com" -Sitecode "AR1" -ApplicationName "M365 Apps for Enterprise" -Path "\\memcm01\source\Applications\Microsoft 365 Apps for Enterprise"

    .NOTES
    FileName:    Update-M365AppsInstallationPackage.ps1
    Author:      Anders Rødland
    Contact:     @AndersRodland
    Created:     2020-12-30
    Version history:
    1.0.0 - (2020-12-30) Script created
#>
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact="Low")]
param(
    [Parameter(HelpMessage='Site server',mandatory=$true)][string]$Siteserver,
    [Parameter(HelpMessage='Site code',mandatory=$true)][string]$Sitecode,
    [Parameter(HelpMessage='Application name of M365 Apps application',mandatory=$true)][string]$ApplicationName,
    [Parameter(HelpMessage='Path to M365 Apps installation package',mandatory=$true)][string]$Path
)

# Define folder variables
$folder = "$path\Office"
$backup = "$folder.bak"

# Download files
Set-Location $path

Write=Host "Renaming old Office folder temporarly in case we need rollback."
Move-Item -Path $folder -Destination $backup
Write-Host "Downloading latest Office files."
.\setup.exe /download "$path/configuration.xml"

# Assume failure unless setup process created new folder structure.
$success = $false

# Verify that files downloaded
if (Test-Path $folder) {
    # Update successful
    $success = $true
}
else {
    # Update failed. Rollback
    Write-Host "Something went wrong. Performing rollback."
    Move-Item -Path $backup -Destination $folder
    $success = $false
}

# We only update distribution points in MEMCM if update of files was successful
if ($success -eq $true) {
    Write-Host "Removing temporary backup folder."
    Remove-Item -Path $backup -Force -Recurse

    # Connect to MEMCM
    Write-Host "Loading Microsoft Endpoint Configuration Manager PowerShell module."
    if((Get-Module ConfigurationManager) -eq $null) { Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" }
    
    # Connect to the site's drive if it is not already present
    Write-Host "Connecting to siteserver $siteserver with sitecode $sitecode"
    if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) { New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer }
    Set-Location "$($SiteCode):\"

    # Update distribution points
    Write-Host "Updating distribution points for $ApplicationName."
    $DeploymentTypeName = (Get-CMDeploymentType -ApplicationName $ApplicationName).LocalizedDisplayName
    Update-CMDistributionPoint -ApplicationName $ApplicationName -DeploymentTypeName $DeploymentTypeName
}
