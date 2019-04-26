<#
.SYNOPSIS
These Functions Are Used Azure AD Management
#>

Function Connect-AzureADPSSession {
    Param(
        [switch]$Local
    )
    if ($Local) {
        Connect-AzureAD
    } else {
        try {
            Connect-AzureAD -Credential $Cred
            Write-Host "INFO: Successfully Connected To Azure AD Module"
        } catch {
            Write-Host "ERROR: Failed To Connected To Azure AD Module"
            Write-Host "$($_)"
            Exit 1
        }
    }
}

Function Get-AzureADUserInfo {
    param (
        [Parameter(Mandatory=$true)][string]$UserName
    )
    Return Get-AzureADUser -SearchString $UserName -All $True
}

Function Get-AzureADOwnerOf {
    param (
        [Parameter(Mandatory=$true)][string]$UserObjectID
    )
    Return Get-AzureADUserOwnedObject -ObjectID $UserObjectID
}

Function Get-AzureADMemberOf {
    param (
        [Parameter(Mandatory=$true)][string]$UserObjectID
    )
    Return Get-AzureADUserMembership -ObjectID $UserObjectID
}

Function Remove-AzureADMemberOf {
    param (
        [Parameter(Mandatory=$true)]$User,
        [Parameter(Mandatory=$true)]$Group
    )
    try {
        Remove-AzureADGroupMember -ObjectId $Group.ObjectID -MemberId $User.ObjectID
        Write-Host "INFO: Removed $($User.DisplayName) As A Member From $($Group.Displayname)"
    } catch {
        Write-Host "ERROR: Removing $($User.DisplayName) As A Member From $($Group.Displayname)"
        Write-Host "$($_)"
    }
}

Function Remove-AzureADOwnerOf {
    param (
        [Parameter(Mandatory=$true)]$User,
        [Parameter(Mandatory=$true)]$Group
    )
    $CurrentOwners = Get-AzureADOwner -Group $Group
    if ($CurrentOwners.count -eq 1) {
        $UserInfo = Get-AzureADUserInfo -UserName "AUTOMATION_ACCOUNT"
        Add-AzureADOwnerOf -User $UserInfo -Group $Group
    }
    try {
        Remove-AzureADGroupOwner -ObjectId $Group.ObjectID -OwnerId $User.ObjectID
        Write-Host "INFO: Removed $($User.DisplayName) As An Owner From $($Group.Displayname)"
    } catch {
        Write-Host "ERROR: Removing $($User.DisplayName) As An Owner From $($Group.Displayname)"
        Write-Host "$($_)"
    }
}

Function Add-AzureADOwnerOf {
    param (
        [Parameter(Mandatory=$true)]$User,
        [Parameter(Mandatory=$true)]$Group
    )
    try {
        Add-AzureADGroupOwner -ObjectId $Group.ObjectID -RefObjectId $User.ObjectID
        Write-Host "INFO: Added $($User.DisplayName) As An Owner To $($Group.Displayname)"
    } catch {
        Write-Host "ERROR: Adding $($User.DisplayName) As An Owner To $($Group.Displayname)"
        Write-Host "$($_)"
    }
}

Function Get-AzureADOwner {
    param (
        [Parameter(Mandatory=$true)]$Group
    )
    try {
        return Get-AzureADGroupOwner -ObjectId $Group.ObjectID
    } catch {
        Write-Host "ERROR: Getting Group Owners For $($Group.Displayname)"
        Write-Host "$($_)"
    }
}

Function Get-ADConnSyncStats {
    $DefaultError = $ErrorActionPreference
    $Global:ErrorActionPreference = "Stop"
    $Cred = Get-AzureADAutomationCred -DictKey "conn_sync"
    try {
        $SyncStats = Invoke-Command -ComputerName $Cred.Server -ScriptBlock {Get-ADSyncScheduler} -Credential $Cred.Credential
        $Global:ErrorActionPreference = $DefaultError
        Write-Host "INFO: Got AD Conn Sync Stats"
        Return $SyncStats
    } catch {
        Write-Host "ERROR: Getting AD Sync Statistics"
        Write-Host "$($_)"
        $Global:ErrorActionPreference = $DefaultError
        Exit 1
    }
}

Function Get-ADConnSyncStatus {
    # Get Sync Stats To See If A Sync Is Running Or Will Run Soon
    $SyncStats = Get-ADConnSyncStats
    if ($SyncStats.SyncCycleInProgress -eq $True) {
        Write-Host "INFO: AD Conn Delta Sync Is In Progress - Sleeping For 90 Seconds"
        Start-Sleep -s 90
        # Get Sync Stats Again To Get New NextSyncCycleStartTimeInUTC
        $SyncStats = Get-ADConnSyncStats
    }
    # Sleep Until 5 Minutes Before Next Sync
    $NextSyncUTC = $SyncStats.NextSyncCycleStartTimeInUTC
    $BeforeNextSync = Get-Date($NextSyncUTC).AddMinutes(-5)
    while ($TimeUTC -lt $BeforeNextSync) {
        $TimeUTC = [System.DateTime]::UtcNow
        Start-Sleep -s 10
    }
}

Function Get-CompanyLastDirSyncTime {
    $DefaultError = $ErrorActionPreference
    $Global:ErrorActionPreference = "Stop"
    try {
        $CompanyLastDirSyncTime = Get-AzureADTenantDetail | Select CompanyLastDirSyncTime
        $Global:ErrorActionPreference = $DefaultError
        Write-Host "INFO: Got CompanyLastDirSyncTime - $($CompanyLastDirSyncTime.CompanyLastDirSyncTime)"
        Return $CompanyLastDirSyncTime
    } catch {
        Write-Host "ERROR: Getting CompanyLastDirSyncTime"
        Write-Host "$($_)"
        $Global:ErrorActionPreference = $DefaultError
        Exit 1
    }
}

Function Get-CompanyLastDirSyncTimeStatus {
    # Get Sync Stats To See When The Last Sync Ran
    $SyncStats = Get-CompanyLastDirSyncTime
    $LastSync = $SyncStats.CompanyLastDirSyncTime
    # The Sync Generally Runs Every +/- 30 Minutes - (28-32 Minute Range After To Be Safe)
    # Add 22 Minutes To Last Sync Since 
    $BeforeNextSync = Get-Date($LastSync).AddMinutes(22)
    while ($TimeUTC -lt $BeforeNextSync) {
        $TimeUTC = [System.DateTime]::UtcNow
        Start-Sleep -s 10
    }
}
