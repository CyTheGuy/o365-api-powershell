<#
.SYNOPSIS
These Functions Are Used O365 Management
#>


Function Connect-o365ExoPSSession {
    param (
        [switch]$UnifiedOnly,
        [switch]$Local
    )
    if ($Local) {
        $Cred = Get-Credential
    } else {
        $Cred = Get-o365AutomationCred
    }
    if ($UnifiedOnly) {
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $Cred -AllowRedirection -Authentication Basic
        Import-Module (Import-PSSession $Session -AllowClobber -CommandName "*-unified*") -Global
        Write-Host "INFO: Successfully Imported Online Exchange Module - Unified Commands Only"
    } else {
        try {
            $Connection = Connect-EXOPSSession -Credential $Cred
            Write-Host "INFO: Successfully Imported Online Exchange Module"
            Import-Module $Connection -Global
        } catch {
            Write-Host "ERROR: Failed To Import Online Exchange Module"
            Write-Host "$($_)"
        }
    }
}

function Enable-o365RemotePowershell {
    param (
        [Parameter(Mandatory=$true)][string]$User,
        [switch]$Enabled
    )
    if ($Enabled) {
        try {
            Set-User -Identity $User -RemotePowerShellEnabled $True
            Write-Host "INFO: Set RemotePowershell = True For $User"
        } catch {
            Write-Host "ERROR: Setting RemotePowershell = True For $User"
            Write-Host "$($_)"
        }
    } else {
        try {
            Set-User -Identity $User -RemotePowerShellEnabled $False
            Write-Host "INFO: Set RemotePowershell = False For $User"
        } catch {
            Write-Host "ERROR: Setting RemotePowershell = False For $User"
            Write-Host "$($_)"
        }
    }
}

function Enable-o365MailboxAuditing {
    param (
        [Parameter(Mandatory=$true)][string]$User,
        [switch]$Enabled
    )
    if ($Enabled) {
        try {
            Set-Mailbox -Identity $User -AuditEnabled $True -AuditOwner Create,HardDelete,MailboxLogin,Move,MoveToDeletedItems,SoftDelete,Update -AuditAdmin Copy,Create,FolderBind,HardDelete,MessageBind,Move,MoveToDeletedItems,SendAs,SendOnBehalf,SoftDelete,Update -AuditDelegate Create,FolderBind,HardDelete,Move,MoveToDeletedItems,SendAs,SendOnBehalf,SoftDelete,Update
            Write-Host "INFO: Added Mailbox Auditing for $User"
        } catch {
            Write-Host "ERROR: Adding Mailbox Auditing For $User"
            Write-Host "$($_)"
        }
    } else {
        try {
            Set-Mailbox -Identity $User -AuditEnabled $False -AuditOwner None
            Write-Host "INFO: Removed Mailbox Auditing for $User"
        } catch {
            Write-Host "ERROR: Removing Mailbox Auditing For $User"
            Write-Host "$($_)"
        }
    }
}

function Enable-o365Addin {
    param (
        [Parameter(Mandatory=$true)][string]$Name,
        [Parameter(Mandatory=$true)][string]$Identity,
        [Parameter(Mandatory=$true)][string]$ProvidedTo,
        [Parameter(Mandatory=$true)][string]$DefaultStateForUser,
        [Parameter(Mandatory=$true)]$UserList
    )
    try {
        Set-App -OrganizationApp -Identity $Identity -ProvidedTo $ProvidedTo -DefaultStateForUser $DefaultStateForUser -UserList $UserList
        Write-Host "INFO: Configured Settings For $Name - Set: ProvidedTo = $ProvidedTo DefaultStateForUser = $DefaultStateForUser UserList = $UserList"
    } catch {
        Write-Host "ERROR: Configuring Settings For $Name"
        Write-Host "$($_)"
    }
}

function Get-o365AllowGuests {
    Write-Host "INFO: Getting Org Config Info For Allow Guests"
    $OrgConfig = Get-OrganizationConfig
    if ($OrgConfig.AllowToAddGuests -eq $True -or $OrgConfig.GuestsEnabled -eq $True) {
        Write-Host "INFO: Guest Access Is Enabled For $($OrgConfig.Identity)  This Needs To Be Disabled By A Global Admin"
        Return $OrgConfig.Identity
    }
}

function Get-o365AllowO365GroupCreation {
    Write-Host "INFO: Getting Org Config Info For Allow O365 Group Creation"
    $OrgConfig = Get-OrganizationConfig
    # GroupsCreationEnabled Determines If People Outside Of o365-group-creation-sg Can Create O365 Groups
    # 47c50fd3-717f-4219-a4c8-413648ba8cca Is The Object ID Of o365-group-creation-sg
    if ($OrgConfig.GroupsCreationEnabled -eq $True -or $OrgConfig.GroupsCreationWhitelistedId -ne "47c50fd3-717f-4219-a4c8-413648ba8cca") {
        Write-Host "INFO: O365 Group Creation Is Enabled For Unapproved Users In $($OrgConfig.Identity) - This Needs To Be Disabled By A Global Admin"
        Return $OrgConfig.Identity
    }
}
function Get-o365DisableAutoForwarding {
    $Domains = @()
    Write-Host "INFO: Getting Remote Domain Info To Check If Auto Forwarding Is Enabled"
    $RemoteDomainSettings = Get-RemoteDomain
    foreach ($Domain in $RemoteDomainSettings) {
        if ($Domain.AutoForwardEnabled -eq $True) {
            Write-Host "INFO: AutoForwarding Is Enabled For $($Domain.Identity) - This Needs To Be Disabled By A Global Admin"
            $Domains += $Domain.Identity
        }
    }
    Return $Domains
}

function Get-o365DisableDirectReportGroup {
    Write-Host "INFO: Getting Org Config Info For Direct Report O365 Groups"
    $OrgConfig = Get-OrganizationConfig
    # DirectReportsGroupAutoCreationEnabled Determines Whether To Enable Or Disable The Automatic Creation Of Direct Report O365 Groups
    if ($OrgConfig.DirectReportsGroupAutoCreationEnabled -eq $True) {

        Write-Host "INFO: Automatic Direct Report O365 Group Creation Is Enabled In - $($OrgConfig.Identity) - This Needs To Be Disabled By A Global Admin"
        Return $OrgConfig.Identity
    }
}

Function Get-o365GroupsGuestsConfig {
    param (
        [switch]$Allow,
        [switch]$Impls
    )
    if ($Allow) {
        if ($Impls) {
            Return Get-UnifiedGroup -Filter {Alias -like "impl-*" -or Alias -like "ext-impl-*"}
        } else {
            Return Get-UnifiedGroup -ResultSize Unlimited | ?{$_.AllowAddGuests -eq $True}
        }
    } else {
        Return Get-UnifiedGroup -ResultSize Unlimited | ?{$_.AllowAddGuests -eq $False}
    }
}

Function Update-o365GroupUser {
    Param(
        [Parameter(Mandatory=$true)][string]$GroupName,
        $Members,
        $Owners
    )
    $DefaultError = $ErrorActionPreference
    $Global:ErrorActionPreference = "Stop"
    # Validate Group Exists
    try {
        $GroupInfo = Get-UnifiedGroup -Identity $GroupName
    } catch {
        Write-Host "ERROR: Cannot Find $GroupName In O365"
        Write-Host "$($_)"
        Return
    }
    # Get Current Members
    $CurrentMembers = Get-UnifiedGroupLinks -Identity $GroupName -LinkType Members
    # Set To Empty Array For Compare If It Is Null Or Add All Owners To Members Since You Need To Be A Member If You Are An Owner
    if ($null -eq $Members) {
        $AllMembers = @()
    } else {
        $AllMembers = @()
        foreach ($User in $Members) {
            $AllMembers += $User.MailNickName
        }
        foreach ($User in $Owners) {
            $AllMembers += $User.MailNickName
        }
        $AllMembers = $AllMembers | Select -Unique
    }
    # Get Current Owners
    $CurrentOwners = Get-UnifiedGroupLinks -Identity $GroupName -LinkType Owners
    # Set To Empty Array For Compare If It Is Null
    if ($null -eq $Owners) {
        $AllOwners = @()
    } else {
        $AllOwners = @()
        foreach ($User in $Owners) {
            $AllOwners += $User.MailNickName
        }
    }
    # Compare Current vs Expected - O365 Member Aliases vs On-premise Member MailNickName 
    $MemberDiff = Compare-Object $CurrentMembers.Alias $AllMembers
    $OwnerDiff = Compare-Object $CurrentOwners.Alias  $AllOwners
    # If You Are An Owner You Need To Be A Member & If You Are No Longer A Member You Need To Be Removed From Owners First If You Are There
    $OwnersToAdd = $($OwnerDiff | ?{$_.SideIndicator -eq '=>'}).InputObject
    if ($OwnersToAdd) {
        try {
            Add-UnifiedGroupLinks -Identity $GroupName -LinkType Members -Links $OwnersToAdd
            Add-UnifiedGroupLinks -Identity $GroupName -LinkType Owners -Links $OwnersToAdd
            Write-Host "INFO: Added $OwnersToAdd As An Owner For $GroupName"
        } catch {
            Write-Host "ERROR: Adding $OwnersToAdd As An Owner For $GroupName"
            Write-Host "$($_)"
        }
    }
    $OwnersToRemove = $($OwnerDiff | ?{$_.SideIndicator -eq '<='}).InputObject
    if ($OwnersToRemove) {
        try {
            Remove-UnifiedGroupLinks -Identity $GroupName -LinkType Owners -Links $OwnersToRemove -Confirm:$False
            Write-Host "INFO: Removed $OwnersToRemove As An Owner From $GroupName"
            Start-Sleep -s 5
        } catch {
            Write-Host "ERROR: Removing $OwnersToRemove As An Owner From $GroupName"
            Write-Host "$($_)"
        }
    }
    $MembersToAdd = $($MemberDiff | ?{$_.SideIndicator -eq '=>'}).InputObject
    if ($MembersToAdd) {
        try {
            Add-UnifiedGroupLinks -Identity $GroupName -LinkType Members -Links $MembersToAdd
            Write-Host "INFO: Added $MembersToAdd As A Member To $GroupName"
        } catch {
            Write-Host "ERROR: Adding $MembersToAdd As A Member To $GroupName"
            Write-Host "$($_)"
        }
    }
    $MembersToRemove = $($MemberDiff | ?{$_.SideIndicator -eq '<='}).InputObject
    if ($MembersToRemove) {
        try {
            Remove-UnifiedGroupLinks -Identity $GroupName -LinkType Members -Links $MembersToRemove -Confirm:$False
            Write-Host "INFO: Removed $MembersToRemove As A Member From $GroupName"
            Start-Sleep -s 5
        } catch {
            Write-Host "ERROR: Removing $MembersToRemove As A Member From $GroupName"
            Write-Host "$($_)"
        }
    }
    $Global:ErrorActionPreference = $DefaultError
}

Function Set-o365GroupConfig {
    param (
        [Parameter(Mandatory=$true)][string]$GroupName,
        [bool]$UnifiedGroupWelcomeMessageEnabled = $False,
        [bool]$HiddenFromAddressListsEnabled = $False
    )
    try {
        Set-UnifiedGroup -Identity $GroupName -UnifiedGroupWelcomeMessageEnabled:$UnifiedGroupWelcomeMessageEnabled -HiddenFromAddressListsEnabled:$HiddenFromAddressListsEnabled
        Write-Host "INFO: Configured $GroupName So Hidden From The GAL = $HiddenFromAddressListsEnabled And Welcome Messages Enabled = $UnifiedGroupWelcomeMessageEnabled"
    } catch {
        Write-Host "ERROR: Configuring $GroupName So Hidden From The GAL = $HiddenFromAddressListsEnabled And Welcome Messages Enabled = $UnifiedGroupWelcomeMessageEnabled"
        Write-Host "$($_)"
    }
}

Function New-o365Group {
    param(
        [Parameter(Mandatory=$true)]$GroupJson,
        $OnPremiseInfo
    )

    $GroupParams = @{}
    # Things Passed In From On-premise
    #
    if ($OnPremiseInfo.RequireSenderAuthenticationEnabled) {
        $GroupParams['RequireSenderAuthenticationEnabled'] = $OnPremiseInfo.RequireSenderAuthenticationEnabled
    }
    # Set The Group Privacy - You Can Change The Privacy Type At Any Point In The Lifecycle Of The Group
    if (($OnPremiseInfo) -and ($OnpremiseInfo.MemberJoinRestriction -notin @('ApprovalRequired', 'Closed'))) {
        $GroupParams['AccessType'] = 'Public'
    } else { 
        $GroupParams['AccessType'] = 'Private'
    }

    # Things Passed In Info From The Json File
    #
    # HiddenGroupMembershipEnabled - You Can't Change This Setting After You Create The Group - This Is A Switch
    if ($GroupJson.HiddenGroupMembershipEnabled) {
        $GroupParams['HiddenGroupMembershipEnabled'] = [System.Convert]::ToBoolean($GroupJson.HiddenGroupMembershipEnabled)
        # Only A Private Group Can Be Marked With Hidden Membership
        $GroupParams['AccessType'] = 'Private'
    }
    if ($GroupJson.NewGroup) {
        $GroupParams['Alias']       = $GroupJson.NewGroup
        $GroupParams['DisplayName'] = $GroupJson.NewGroup
    } elseif ($GroupJson.NewName) {
        # Change In Name There Is No Need To Keep All The Proxy Addresses
        $GroupParams['Alias']       = $GroupJson.NewName
        $GroupParams['DisplayName'] = $GroupJson.NewName
    } else {
        $GroupParams['Alias']       = $OnPremiseInfo.Alias
        $GroupParams['DisplayName'] = $OnPremiseInfo.DisplayName
        # Keep All Proxy Addresses Since You Kept The Same Name
        $GroupParams['EmailAddresses'] = $OnPremiseInfo.EmailAddresses
    }
    if ($GroupParams.Alias.Length -gt 64) {
        $Alias = $GroupParams.Alias
        $GroupParams['Alias'] = ($Alias.Substring(0,62) + (Get-Random -Minimum 10 -Maximum 99))
        Write-Host "WARNING: $Alias Was Greater Than 64 Characters So It Has Been Trimmed To - $($GroupParams.Alias)"
    }
    $ExistsAlias = Get-UnifiedGroup -Identity $GroupParams.Alias -ErrorAction 'SilentlyContinue'
    $ExistsDisplayName = Get-UnifiedGroup -Identity $GroupParams.DisplayName -ErrorAction 'SilentlyContinue'
    if (!$ExistsAlias -and !$ExistsDisplayName) {
        try {
            if ($GroupJson.AutoSubscribeNewMembers) {
                New-UnifiedGroup @GroupParams -AutoSubscribeNewMembers -SubscriptionEnabled
                Write-Host "INFO: Created $($GroupParams.Alias) With Auto Subscribe New Members"
            } else {
                New-UnifiedGroup @GroupParams -SubscriptionEnabled
                Write-Host "INFO: Created $($GroupParams.Alias) Without Auto Subscribe New Members"
            }
            if ($OnPremiseInfo.HiddenFromAddressListsEnabled -eq $True) {
                Write-Host "INFO: Group Was Hidden From GAL On-Premise Hiding It From GAL Again"
                Set-o365GroupConfig -GroupName $GroupParams.Alias -HiddenFromAddressListsEnabled $True
            } else {
                Set-o365GroupConfig -GroupName $GroupParams.Alias
            }
        } catch {
            Write-Host "ERROR: Creating $($GroupParams.Alias) With The Following Exception:"
            Write-Host "$($_)"
            Return 1
        }
    } else {
        Write-Host "WARNING: There Is Already An O365 Group Named $($GroupParams.Alias) Or $($GroupParams.DisplayName)"
        Return 2
    }
}

Function New-o365GroupNotOnprem {
    param(
        [Parameter(Mandatory=$true)][string]$GroupName
    )

    $GroupParams = @{}
    if ($GroupName.Length -gt 64) {
        $GroupParams['DisplayName'] = $GroupName
        $GroupParams['Alias'] = ($GroupName.Substring(0,62) + (Get-Random -Minimum 10 -Maximum 99))
        Write-Host "WARNING: $GroupName Was Greater Than 64 Characters So It Has Been Trimmed To - $($GroupParams.Alias)"
    } else {
        $GroupParams['DisplayName'] = $GroupName
        $GroupParams['Alias'] = $GroupName
    }
    $ExistsAlias = Get-UnifiedGroup -Identity $GroupParams.Alias -ErrorAction 'SilentlyContinue'
    $ExistsDisplayName = Get-UnifiedGroup -Identity $GroupParams.DisplayName -ErrorAction 'SilentlyContinue'
    if (!$ExistsAlias -and !$ExistsDisplayName) {
        try {
            New-UnifiedGroup @GroupParams -SubscriptionEnabled
            Write-Host "INFO: Created $($GroupParams.Alias) Without Auto Subscribe New Members Turned On"
            Set-o365GroupConfig -GroupName $GroupParams.Alias -HiddenFromAddressListsEnabled $True
        } catch {
            Write-Host "ERROR: Creating $($GroupParams.Alias) With The Following Exception:"
            Write-Host "$($_)"
            Return 1
        }
    } else {
        Write-Host "WARNING: There Is Already An O365 Group Named $($GroupParams.Alias) Or $($GroupParams.DisplayName)"
        Return 2
    }
}

Function Rename-O365Group {
    Param(
        [Parameter(Mandatory=$true)][string]$GroupName,
        [Parameter(Mandatory=$true)][string]$NewName,
        [switch]$KeepOldAliases
    )

    $DefaultError = $ErrorActionPreference
    $Global:ErrorActionPreference = "Stop"

    $PrimarySmtpAddress = $NewName + '@palantir.com'
    $OtherEmail = $NewName + '@palantirtech.onmicrosoft.com'

    $GroupInfo = Get-UnifiedGroup -Identity $GroupName -ErrorAction 'SilentlyContinue'
    if ($GroupInfo) {
        if ($KeepOldAliases) {
            try {
                # You can't use the EmailAddresses and PrimarySmtpAddress parameters in the same command
                Set-UnifiedGroup -Identity $GroupName -Alias $NewName -DisplayName $NewName -PrimarySmtpAddress $PrimarySmtpAddress
                Start-Sleep -s 3
                Set-UnifiedGroup -Identity $NewName -EmailAddresses @{Add="smtp:$OtherEmail"}
                Write-Host "INFO: Changed $GroupName To $NewName And Kept The Old Aliases"
            } catch {
                Write-Host "ERROR: Renaming $GroupName To $NewName"
                Write-Host "$($_)"
            }
        } else {
            try {
                Set-UnifiedGroup -Identity $GroupName -Alias $NewName -DisplayName $NewName -EmailAddresses "SMTP:$PrimarySmtpAddress","smtp:$OtherEmail"
                Write-Host "INFO: Changed $GroupName To $NewName And Removed The Old Aliases"
            } catch {
                Write-Host "ERROR: Renaming $GroupName To $NewName"
                Write-Host "$($_)"
            }
        }
    } else {
        Write-Host "WARNING: $GroupName Does Not Exist In O365"
    }
    $Global:ErrorActionPreference = $DefaultError 
}
