<#
.SYNOPSIS
    Create a set of usefull Software Updates Collections for ConfigMgr Current Branch

.DESCRIPTION
    This script will create a set of predefined Software Updates Collections that may be useful to ConfigMgr administrators.
    Client collections are stored in the folder "Software Updates\Client Updates"
    Server collections are stored in the folder "Software Updates\Server Updates"
    

.PARAMETER SiteServer
    Site server name with SMS Provider installed

.PARAMETER ClientWaves
    Amount of "Client Updates - Wave X" collections to create. Optional parameter with defualt value 4.

.PARAMETER ServerWaves
    Amount of "Server Updates - Wave X" collections to create Optional parameter with defualt value 2.

.EXAMPLE
    New-SoftwareUpdatesCollections -SiteServer CM01 -ClientWaves 4 -ServerWaves 2

.NOTES
    Script name: New-SoftwareUpdatesCollections.ps1
    Author:      Anders Rødland
    Contact:     @AndersRodland
    DateCreated: 2018-01
#>

[CmdLetBinding()]
Param(
    [Parameter(Mandatory=$True, HelpMessage='Configuration Manager Site Server')][String]$SiteServer,
    [Parameter(Mandatory=$False, HelpMessage='Configuration Manager Site Server')][String]$ClientWaves=4,
    [Parameter(Mandatory=$False, HelpMessage='Configuration Manager Site Server')][String]$ServerWaves=2
    )

Begin {
    # Determine SiteCode from WMI
    Try {
        Write-Verbose "Determining SiteCode for Site Server: '$($SiteServer)'"
        $SiteCodeObjects = Get-WmiObject -Namespace "root\SMS" -Class SMS_ProviderLocation -ComputerName $SiteServer -ErrorAction Stop
        foreach ($SiteCodeObject in $SiteCodeObjects) {
            if ($SiteCodeObject.ProviderForLocalSite -eq $true) {
                $SiteCode = $SiteCodeObject.SiteCode
                Write-Debug "SiteCode: $($SiteCode)"
            }
        }
    }
    Catch [Exception] { Write-Warning -Message "Unable to determine SiteCode" ; Break }
    
    # Import the ConfigurationManager.psd1 module 
    if((Get-Module ConfigurationManager) -eq $null) {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
    }
    
    # Connect to the site's drive if it is not already present
    if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer
    }
    
    # Store the current location
    $MyCurrentLocation = Get-Location
    $SiteCode = $SiteCode + ":"
    # Set the current location to be the site code.
    Try { Set-Location $SiteCode }
    Catch [Exception] { Write-Warning -Message "Error, could not set location to $SiteCode" ; Break }
}

Process {
    $ClientCollections = "Client Updates - Pre Pilot", "Client Updates - Pilot"
    $ServerCollections = "Server Updates - Pre Pilot", "Server Updates - Pilot"
    
    $ClientLimitingCollection = "All Systems"
    $ServerLimitingCollection = "All Systems"
    
    $RootFolder = "Software Updates"
    $ClientFolder = "Client Updates"
    $ServerFolder = "Server Updates"
    
    if (-not (Test-Path "$SiteCode\Devicecollection\$RootFolder")) {New-Item -Path "$SiteCode\Devicecollection" -Name $RootFolder }
    else { Write-Output "Folder '$SiteCode\Devicecollection\$RootFolder' already exist. Skipping..." }

    if (-not (Test-Path "$SiteCode\Devicecollection\$RootFolder\$ClientFolder")) { New-Item -Path "$SiteCode\Devicecollection\$RootFolder" -Name $ClientFolder }
    else { Write-Output "Folder '$SiteCode\Devicecollection\$RootFolder\$ClientFolder' already exist. Skipping..." }

    if (-not (Test-Path "$SiteCode\Devicecollection\$RootFolder\$ServerFolder")) { New-Item -Path "$SiteCode\Devicecollection\$RootFolder" -Name $ServerFolder }
    else { Write-Output "Folder '$SiteCode\Devicecollection\$RootFolder\$ServerFolder' already exist. Skipping..." }
    
    for ($i = 1; $i -le $ClientWaves; $i++) {
        $ClientCollections += "Client Updates - Wave $i"
    }
    
    foreach ($name in $ClientCollections) {
        if (-not (Get-CMDeviceCollection -Name $name)) {
            try {
                New-CMDeviceCollection -Name $name -LimitingCollectionName $ClientLimitingCollection | Out-Null
                Write-Output "Collection '$name' created successfully."
                $InputObject = Get-CMDeviceCollection -Name $name
                Move-CMObject -InputObject $InputObject -FolderPath "$site\Devicecollection\$RootFolder\$ClientFolder"
                Write-Verbose "Collection '$name' moved successfully to '$site\Devicecollection\$RootFolder\$ClientFolder'"
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
                Write-Output "Error creating client collection. Errormessage: $ErrorMessage. Failed item: $FailedItem."
            }
        }
        else { Write-Output "Collection '$name' already exist. Skipping..." }
    }
    
    
    for ($i = 1; $i -le $ServerWaves; $i++) {
        $ServerCollections += "Server Updates - Wave $i"
    }
    
    foreach ($name in $ServerCollections) {
        if (-not (Get-CMDeviceCollection -Name $name)) {
            try {
                New-CMDeviceCollection -Name $name -LimitingCollectionName $ServerLimitingCollection | Out-Null
                Write-Output "Collection '$name' created successfully."
                $InputObject = Get-CMDeviceCollection -Name $name
                Move-CMObject -InputObject $InputObject -FolderPath "$site\Devicecollection\$RootFolder\$ServerFolder"
                Write-Verbose "Collection '$name' moved successfully to '$site\Devicecollection\$RootFolder\$ServerFolder'"
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
                Write-Output "Error creating client collection. Errormessage: $ErrorMessage. Failed item: $FailedItem."
            }
            
        }
        else { Write-Output "Collection '$name' already exist. Skipping..." }
    }
}

End {
    # Set-Location to previsou location before changing to SiteCode:
    Set-Location $MyCurrentLocation
}
