#####################################
# Azure AD Module Functions
#####################################

Function Update-AzureADModules {
    
    $AzureModules = Get-Module -ListAvailable AzureAD*

    foreach ($Module in $AzureModules.Name) {

        Uninstall-Module $Module
        Write-Host "Uninstalled: Module $Module"
    }

    Install-Module AzureADPreview
    Write-Host "Installed: Lastest AzureADPreview Module"
}

Function Import-AzureADModules {
    
    $LoadedModules = Get-Module 

    if ($LoadedModules.Name -notcontains "AzureADPreview") {

        Import-Module AzureADPreview
        Write-Host "Imported: AzureADPreview Module"
    }
}

####################################
# AD Specific Functions
####################################

function Get-ADUserMembers {

  param (
    [Parameter(Mandatory=$false)][string]$GroupString,
    [Parameter(Mandatory=$false)][array]$Groups
  )

  $AllGroupsMembers = @{}

  if ($GroupString) {
    if ($Groupstring -like "*") {
      $AllGroups = Get-ADGroup -Filter {MailNickName -like $GroupString} -Properties MailNickName
    }
    else {
      $AllGroups = Get-ADGroup -Filter {MailNickName -eq $GroupString} -Properties MailNickName
    }
  }
  else {
    $AllGroups = $Groups
  }

  foreach ($Group in $AllGroups) {

    $MembersSAMs = Get-ADGroupMember -Identity $Group.SAMAccountName -Recursive | Select SAMAccountName
    $MembersAlias = $MembersSAMs | %{Get-ADUser -Identity $_.SAMAccountName -Properties MailNickName} | Select MailNickName

    if ($Group.Alias) {
      $AllGroupsMembers.Add($Group.Alias, $MembersAlias.MailNickName)
    }
    else {
      $AllGroupsMembers.Add($Group.MailNickName, $MembersAlias.MailNickName)
    }
  }

  # Add Group mailNickName & Members SAMAccountName To Array
  $AllGroupsMembers.Keys | Select @{l='Alias';e={$_}},@{l='NewAlias';e={"o365_"+$($_)}},@{l='Members';e={$AllGroupsMembers.$_}}
  Return $AllGroupsMembers
}

function Get-ADGroupDescriptions {

  param (
    [Parameter(Mandatory=$false)][string]$GroupString,
    [Parameter(Mandatory=$false)][array]$Groups
  )

  $AllGroupsDescriptions = @()

  if ($GroupString) {
    $Groups = Get-ADGroup -Filter {SAMAccountName -like $GroupString} 
  }

  foreach ($Group in $Groups) {
    $Descriptions = Get-ADGroup -Identity $Group.SAMAccountName -Properties MailNickName,Description | Select @{l='Alias';e={$_.MailNickName}},@{l='NewAlias';e={"o365_"+$($_.MailNickName)}},SAMAccountName,Description,@{l='Notes';e={$($_.SAMAccountName) + " - " + $($_.Description)}}
    $AllGroupsDescriptions += $Descriptions
  }

  Return $AllGroupsDescriptions
}

####################################
# Exchange Specific Functions
####################################

function Get-ExGroupInfo {

  param (
    [Parameter(Mandatory=$false)][string]$GroupString,
    [Parameter(Mandatory=$false)][string]$Group,
    [Parameter(Mandatory=$false)][array]$Groups
  )

  $AllGroupsInfo = @()
  Add-ExchangeSnapin

  if ($GroupString) {
    $AllGroupsInfo = Get-DistributionGroup -Filter {Alias -like $GroupString} | Select Alias,DisplayName,@{l='NewAlias';e={"o365_"+$($_.Alias)}},@{l='NewDisplayName';e={"o365_"+$($_.DisplayName)}},SAMAccountName,@{l='AccessType';e={$_.MemberJoinRestriction}},RequireSenderAuthenticationEnabled
  }
  if ($Group) {
    $AllGroupsInfo = Get-DistributionGroup -Identity $Group | Select Alias,DisplayName,@{l='NewAlias';e={"o365_"+$($_.Alias)}},@{l='NewDisplayName';e={"o365_"+$($_.DisplayName)}},SAMAccountName,@{l='AccessType';e={$_.MemberJoinRestriction}},RequireSenderAuthenticationEnabled
  }
  if ($Groups) {
    $AllGroupsInfo = $Groups | %{Get-DistributionGroup -Identity $_.Alias} | Select Alias,DisplayName,@{l='NewAlias';e={"o365_"+$($_.Alias)}},@{l='NewDisplayName';e={"o365_"+$($_.DisplayName)}},SAMAccountName,@{l='AccessType';e={$_.MemberJoinRestriction}},RequireSenderAuthenticationEnabled
  }
  
  Return $AllGroupsInfo
}

function Get-ExGroupOwners {

  param (
    [Parameter(Mandatory=$false)][string]$GroupString,
    [Parameter(Mandatory=$false)][array]$Groups
  )

  $AllGroupsOwners = @()
  Add-ExchangeSnapin
  
  if ($GroupString) {
    $AllGroupsOwners = Get-DistributionGroup -Filter {Alias -like $GroupString} | Select-Object Alias,@{l='NewAlias';e={"o365_"+$($_.Alias)}},ManagedBy
  }
  else {
    $AllGroupsOwners =  $Groups | %{Get-DistributionGroup -Identity $_.Alias} | Select-Object Alias,@{l='NewAlias';e={"o365_"+$($_.Alias)}},ManagedBy
  }

  Return $AllGroupsOwners
}

######################################
# O365 Exchange Group Config Functions
######################################
Function Disable-DisableAllowGuestsGlobal {
    
    # This function is used if you want to set it on a global basis for all group or you can set it through the Admin GUI
    # This will require global admin permissions
    # See if you already have an AzureADDirectorySetting object, and if so save the Object ID
    try {
        $SettingsObjectID = (Get-AzureADDirectorySetting | Where-Object -Property Displayname -Value "Group.Unified" -EQ).ID
    }
    catch {
        BasicLogging -Log "$($_)"
        $ExceptionMessage = $Error[0].Exception | Format-List * -Force
        BasicLogging -Log "$ExceptionMessage"
    }

    $SettingsCopy = Get-AzureADDirectorySetting –Id $SettingsObjectID
    $SettingsCopy["AllowGuestsToAccessGroups"] = $False
    $SettingsCopy["AllowGuestsToBeGroupOwner"] = $False
    $SettingsCopy["AllowToAddGuests"] = $False

    Set-AzureADDirectorySetting -Id $SettingsObjectID -DirectorySetting $SettingsCopy
}
Function Disable-AllowGuests {

    # This function is used if you want to set it on a PER group basis
    param (
        [Parameter(Mandatory=$false)][string]$GroupString,
        [Parameter(Mandatory=$false)][array]$Groups    
    )
                                                                                                                                
    $Template = Get-AzureADDirectorySettingTemplate | ?{$_.DisplayName -eq "group.unified.guest"}
    $SettingsCopy = $Template.CreateDirectorySetting()
    $SettingsCopy["AllowToAddGuests"]=$False

    if ($GroupString) {
        $Groups = Get-AzureADGroup -SearchString $GroupString
    }
    else {
        $Groups = $Groups | %{Get-AzureADGroup -SearchString "$($_.Alias)"}
    }

    foreach ($Group in $Groups) {

        $Mail = $Group.Mail
        $GroupID = (Get-AzureADGroup -SearchString $Mail).ObjectId

        try { 
          New-AzureADObjectSetting -TargetType Groups -TargetObjectId $GroupID -DirectorySetting $SettingsCopy 
        }
        catch {
          BasicLogging -Log "ERROR: Disabling Allow Guests For $GroupID"
          BasicLogging -Log "$($_)"
          $ExceptionMessage = $Error[0].Exception | Format-List * -Force
          BasicLogging -Log "$ExceptionMessage"
        }
    }
}
Function Get-EXGroupsGuests {

  param (
        [Parameter(Mandatory=$false)][switch]$Allow
    )
                    
  if ($Allow) {
    Get-UnifiedGroup -ResultSize Unlimited | ?{$_.AllowAddGuests -eq $true}
  }  
  else {
    Get-UnifiedGroup -ResultSize Unlimited | ?{$_.AllowAddGuests -eq $false}
  }                                                                                                                          
}

# This disabled the welcome message for any new members of a groupname and hides it from the GAL
Function Set-ConfigGroups {

  param (
      [Parameter(Mandatory=$false)][string]$GroupString,
      [Parameter(Mandatory=$false)][array]$Groups,
      [Parameter(Mandatory=$false)][switch]$Bulk
  )

  if ($GroupString) {                                                                                                                       
    $Groups = Get-UnifiedGroup -Filter {Alias -like $GroupString}
  }

  if ($Bulk) {
    $Groups.Alias | Set-UnifiedGroup -UnifiedGroupWelcomeMessageEnabled:$false -HiddenFromAddressListsEnabled:$true
    BasicLogging -Log "Configured $Groups so they are hidden from the GAL and they dont have welcome messages"
  }
  else {
    foreach ($Group in $Groups.Alias) {
      Set-UnifiedGroup -Identity $Group -UnifiedGroupWelcomeMessageEnabled:$false -HiddenFromAddressListsEnabled:$true
      BasicLogging -Log "Configured $Group so it is hidden from the GAL and that it doesnt have welcome messages"
    }
  }
}

######################################
# O365 Exchange Functions
######################################

Function Get-O365GroupsAllInfo {

    Param(
        [Parameter(Mandatory=$false)][array]$Groups,
        [Parameter(Mandatory=$false)][string]$GroupString,
        [Parameter(Mandatory=$false)][array]$Attributes
    )

    if ($GroupString) {
        # Return Specific Groups & Specific Attributes
        Return Get-UnifiedGroup -Filter {Alias -like $Groupstring} #### THIS DOESNT WORK LIKE EXPECTED IT GRABS ALL GROUPS AND IS DUMB AND DOESNT ONLY GRAB impl-*
    }
    elseif (($Groups) -and ($Attributes)) {
        # Return Specific Groups & Specific Attributes
        Return $Groups | Get-UnifiedGroup | Format-List $Attributes
    }
    elseif ($Groups) {
        # Return Specific Groups & All Attributes
        Return $Groups | Get-UnifiedGroup | Format-List
    }
    elseif ($Attributes) {
        # Return All Groups & Specific Attributes
        Return Get-UnifiedGroup | Format-List $Attributes
    }
    else {
        # Return All Groups & All Attributes
        Return Get-UnifiedGroup | Format-List
    } 
}

Function Remove-O365Groups {
    Param(
        [Parameter(Mandatory)]$Groups,
        [Parameter(Mandatory=$false)][switch]$Bulk,
        [Parameter(Mandatory=$false)][switch]$SafeMode
    )

    if ($Bulk) {
        if (!$SafeMode) {
            try {
                BasicLogging -Log "$($Groups.Alias) Will Be Destroyed"
                $Groups.Alias | Remove-UnifiedGroup -Confirm:$False
                BasicLogging -Log "Destroyed $($Groups.Alias)"
            } 
            catch {
                BasicLogging -Log "ERROR: Failed To Destroy $Groups"
                Exit 1
            }
        }
        else {
            BasicLogging -Log "$Groups Would Have Been Destroyed"
        }
    }
    else {
        #foreach ($Group in $Groups.Alias) {
        foreach ($Group in $Groups) {
            if (!$SafeMode) {
                try {
                    BasicLogging -Log "$Group Will Be Destroyed"
                    Remove-UnifiedGroup -Identity $Group -Confirm:$False
                    BasicLogging -Log "Destroyed $Group"
                } 
                catch {
                    BasicLogging -Log "ERROR: Failed To Destroy $Group"
                }
            } 
            else {
                BasicLogging -Log "$Group Would Have Been Destroyed"
            }
        }
    }
}

Function Create-O365Groups {
    Param(
        [Parameter(Mandatory)]$Groups,
        [Parameter(Mandatory=$false)][switch]$SafeMode,
        [Parameter(Mandatory=$false)][switch]$Migration
    )

    foreach ($Group in $Groups) {
        if (!$SafeMode) {

            if ($Migration) {
              $Alias = $Group.NewAlias
              $DisplayName = $Group.NewDisplayName
            }
            else {
              $Alias = $Group.Alias
              $DisplayName = $Group.DisplayName
            }

            if ($Alias.length -gt 64) {
                BasicLogging -Log "ERROR: $Alias is greater than 64 characters so its being skipped because its too large of a name"
                Continue
            }
            
            if ($Group.AccessType -eq "ApprovalRequired") {

              $AccessType = "Private"
            }
            elseif ($Group.AccessType -eq "Closed") {

              $AccessType = "Private"
            }
            else {

              $AccessType = "Public"
            }

            $RequireSenderAuthenticationEnabled = $Group.RequireSenderAuthenticationEnabled 
            $RequireSenderAuthenticationEnabled = [System.Convert]::ToBoolean($RequireSenderAuthenticationEnabled)
            
            # When you create an Office 365 Group without using the EmailAddresses parameter, the Alias value you specify is used to generate the primary email address 
            # If you don't use the Alias parameter when you create an Office 365 Group, the value of the DisplayName parameter is used. 
            # DisplayName > Alias > Email Address
            $ExistsAlias = Get-UnifiedGroup -Identity $Alias -ErrorAction 'SilentlyContinue'
            $ExistsDisplayName = Get-UnifiedGroup -Identity $DisplayName -ErrorAction 'SilentlyContinue'

            if (!$ExistsAlias -and !$ExistsDisplayName) {

                try { 
                  BasicLogging -Log "$Alias Will Be Created"
                  New-UnifiedGroup -Alias $Alias -Name $Alias -DisplayName $DisplayName -PrimarySmtpAddress "$Alias@test.com" -RequireSenderAuthenticationEnabled:$RequireSenderAuthenticationEnabled -AccessType $AccessType -AutoSubscribeNewMembers -SubscriptionEnabled
                  Set-UnifiedGroup -Identity $Alias -EmailAddresses @{Add="smtp:$Alias@test.mail.onmicrosoft.com","smtp:$Alias@test.onmicrosoft.com"}
                  BasicLogging -Log "Created $Alias, $DisplayName, $AccessType"
                }
                catch {
                  BasicLogging -Log "ERROR: Creating $Alias With The Following Exception:"
                  BasicLogging -Log "$($_)"
                }
            }
            else {

              BasicLogging -Log "$Alias Already Exists, So It Wasnt Created. Updating the Aliases & Primary SMTP Address"
              Set-UnifiedGroup -Identity $Alias -PrimarySmtpAddress "$Alias@test.com"
              Set-UnifiedGroup -Identity $Alias -EmailAddresses @{Add="smtp:$Alias@test.mail.onmicrosoft.com","smtp:$Alias@test.onmicrosoft.com"}
            }
        }
        else {
            BasicLogging -Log "$Group Would Have Been Created"
        }
    }
}

Function Update-O365GroupOwner {
  Param(
      [Parameter(Mandatory)][array]$Groups
  )

  foreach ($Group in $Groups) {

    $NewMemberOwners = @()
    $NewMembers = @()
    $AllOwners = @()
    $RemoveOwners = @()

    $Alias = $Group.NewAlias

    # Audit Existing Owners Since Add-UnifiedGroupLinks Doesnt Clobber Ownership
    $O365GroupOwners = Get-UnifiedGroupLinks -Identity $Alias -LinkType Owners
    $O365GroupMembers = Get-UnifiedGroupLinks -Identity $Alias -LinkType Members

    foreach ($Owner in $Group.ManagedBy) {

        $Owner = $Owner.Name
        $AllOwners += $Owner

        if ($O365GroupMembers.Name -notcontains $Owner -and $O365GroupMembers.DisplayName -notcontains $Owner) {

            BasicLogging -Log "INFO: Will Add $Owner As A O365Member For $Alias"
            $NewMembers += $Owner
        }
        if ($O365GroupOwners.Name -notcontains $Owner -and $O365GroupOwners.DisplayName -notcontains $Owner) {

            BasicLogging -Log "INFO: Will Add $Owner As A O365Owner To $Alias"
            $NewMemberOwners += $Owner
        }
    }

    if ($NewMembers) {

      try { 
        Add-UnifiedGroupLinks -Identity $Alias -LinkType Members -Links $NewMembers
      }
      catch {
        BasicLogging -Log "ERROR: Adding $NewMembers As O365Members To $Alias With The Following Exception:"
        BasicLogging -Log "$($_)"
      }
    }
    if ($NewMemberOwners) {

      try { 
        Add-UnifiedGroupLinks -Identity $Alias -LinkType Owners -Links $NewMemberOwners
      }
      catch {
        BasicLogging -Log "ERROR: Adding $NewMemberOwners As O365Owners To $Alias With The Following Exception:"
        BasicLogging -Log "$($_)"
      }
    }

    foreach ($O365Owner in $O365GroupOwners) {
        
        if ($AllOwners -notcontains $O365Owner.Name -and $AllOwners -notcontains $O365Owner.DisplayName) {
            
            BasicLogging -Log "INFO: Will Remove $($O365Owner.DisplayName) As An O365Owner From $Alias"
            $RemoveOwners += $O365Owner.Name
        }
    }

     if ($RemoveOwners) {

      try { 
        Remove-UnifiedGroupLinks -Identity $Alias -LinkType Owners -Links $RemoveOwners -Confirm:$False
      }
      catch {
        BasicLogging -Log "ERROR: Removing $RemoveOwners As O365Owners From $Alias With The Following Exception:"
        BasicLogging -Log "$($_)"
      }
    }
  }
}
Function Update-O365GroupMembers {
  Param(
      [Parameter(Mandatory)][array]$Groups
  )

  foreach ($Group in $Groups) { 

    $MembersToRemove = @()
    $MembersToAdd = @()
    $GroupAlias = $Group.NewAlias
    $OnpremiseMembers = $Group.Members
    $OnpremiseMembers = $OnpremiseMembers | ?{$_}

    if (!$OnpremiseMembers) {

      BasicLogging -Log "WARNING: $GroupAlias Has No Members On Premise To Add"
      Continue
    }

    # O365: Users cannot be an owner without being a member
    # Onpremise: Users can be an owner without being a member
    # So you need to make sure that you compare O365 Members against O365 Owners AND OnPremise Members
    $O365GroupMembers = Get-UnifiedGroupLinks -Identity $GroupAlias -LinkType Members
    $O365GroupOwners = Get-UnifiedGroupLinks -Identity $GroupAlias -LinkType Owners

    try {

      $MembersDiffs = Compare-Object $O365GroupMembers.Alias $OnpremiseMembers | Select @{l='Alias';e={$_.InputObject}},SideIndicator
    }
    catch {

      BasicLogging -Log "ERROR: Doing Membership Compare For $GroupAlias With The Following Exception:"
      BasicLogging -Log "$($_)"
    }

    if (!$MembersDiffs) {
      # Continue As Membership Is Good
      Continue
    }

    $MembersDiffs = $MembersDiffs | ?{$_.Alias}

    foreach ($Member in $MembersDiffs) {

      $MemberAlias = $Member.Alias

      if ($O365GroupOwners.Alias -contains $MemberAlias) {
        Continue
      }

      # If In O365GroupMembers But Not AllOnPremiseMembers Remove Them From O365 Group
      if ($Member.SideIndicator -eq "<=") {

        BasicLogging -Log "INFO: Will Remove $($Member.Alias) As An O365Member From $GroupAlias"
        $MembersToRemove += $Member.Alias
      }
      # If Not In O365GroupMembers In AllOnPremiseMembers Add Them To O365 Group
      if ($Member.SideIndicator -eq "=>") {

        BasicLogging -Log "INFO: Will Add $($Member.Alias) As An O365Member For $GroupAlias"
        $MembersToAdd += $Member.Alias
      }
    }

    if ($MembersToAdd) {
      try { 

        Add-UnifiedGroupLinks -Identity $GroupAlias -LinkType Members -Links $MembersToAdd
      }
      catch {

        BasicLogging -Log "ERROR: Adding New O365Members to $GroupAlias With The Following Exception:"
        BasicLogging -Log "$($_)"
      }
    }

    if ($MembersToRemove) {
      try { 

        Remove-UnifiedGroupLinks -Identity $GroupAlias -LinkType Members -Links $MembersToRemove -Confirm:$False
      }
      catch {

        BasicLogging -Log "ERROR: Removing Old O365Members From $GroupAlias With The Following Exception:"
        BasicLogging -Log "$($_)"
      }
    }
  }
}

Function Update-O365GroupDescriptions {
    Param(
        [Parameter(Mandatory)][array]$Groups
    )

    foreach ($Group in $Groups) { 

      $Alias = $Group.NewAlias
      $Description = $Group.Notes
      $Notes = Get-UnifiedGroup -Identity $Alias | Select Notes
      
      if ($Notes.Notes -ne $Description) {

        Set-UnifiedGroup -Identity $Alias -Notes $Description
      }
   }
}

######################################
# Azure AD Functions
######################################

Function Get-AzureADGroupInfo {
    
    Param(
        [Parameter(Mandatory=$false)][string]$GroupString, ## VALIDATE THIS DOESNT HAVE * because search string doesnt accept it
        [Parameter(Mandatory=$false)][switch]$All,
        [Parameter(Mandatory=$false)][string]$ID
    )

    if (($All -and $GroupString) -or ($All -and $ID) -or ($GroupString -and $ID)) {

        Return "Only pass a single paramater: All, GroupString, or ID"
    }
    elseif ($GroupString) {

        $AzureADGroupInfo = Get-AzureADMSGroup -SearchString $GroupString -All $true
    }
    elseif ($All) {

        $AzureADGroupInfo = Get-AzureADMSGroup -All $true
    } 
    elseif ($ID) {

        $AzureADGroupInfo = Get-AzureADMSGroup -Id $ID
    }

    Return $AzureADGroupInfo
}

Function Remove-AzureADGroup {

    Param(
        [Parameter(Mandatory)]$Groups,
        [Parameter(Mandatory=$false)][switch]$Hard,
        [Parameter(Mandatory=$false)][switch]$SafeMode
    )

    foreach ($Group in $Groups) {

        $ID = $Group.ID
        $DisplayName = $Group.DisplayName

        if (!$SafeMode){
            if ($Hard) {

                  Remove-AzureADMSDeletedDirectoryObject -Id $ID
                  BasicLogging -Log "Hard Deleted AzureAD Group: $ID, $DisplayName"
            }
            else {

                Remove-AzureADMSGroup -Id $ID
                BasicLogging -Log "Soft Deleted AzureAD Group: $ID, $DisplayName"
            }
        }
        else {
            BasicLogging -Log "$ID, $DisplayName Would Have Been Destroyed"
        }
    }
}

Function Get-AzureADSoftDeletedGroups {
    
    Write-Host "Getting Soft Deleted Groups That Were Deleted Within 30 Days Ago"
    $SoftDeletedGroups = Get-AzureADMSDeletedGroup -All $True

    if (!$SoftDeletedGroups) {
        Return "There Are No Soft Deleted Groups"
    }
    else {
        Return $SoftDeletedGroups
    }
}

Function Restore-AzureADSoftDeletedGroups {

    Param(
        [Parameter(Mandatory)]$Groups,
        [Parameter(Mandatory=$false)][switch]$SafeMode
    )

    $SoftDeletedGroups = Get-AzureADSoftDeletedGroups

    foreach ($ID in $Groups.ID) {

        if (!($SoftDeletedGroups.ID -contains $ID)) {

            BasicLogging -Log "$ID Is Not A Soft Deleted Group"
        }
        else {
            if (!$SafeMode){
                BasicLogging -Log "Restoring: $ID"
                Restore-AzureADMSDeletedDirectoryObject –Id $ID | Out-Null
            }
            else {
                BasicLogging -Log "$ID Would Have Been Restored"
            }
        }
    }
}

######################################
# Logging
######################################

# For Powershell Below v5
Function BasicLogging {
    
    Param(
        [Parameter(Mandatory=$true)][string]$Log,
        [Parameter(Mandatory=$false)][switch]$Email
    )

    $LogRoot = "C:\Logs\O365Scripts\O365GroupMigration\"
    $FileTimeString = (Get-Date).ToString("yyyy-MM-dd")
    $LogFile = ($LogRoot + "O365_Group_Migration_" + $FileTimeString + ".log")

    Write-Host "$Log"
    Add-Content $LogFile "$Log"
}

Function CheckLogPathExists {

    Param(
        [Parameter(Mandatory=$true)][string]$LogPath
    )

    $PathExists = Test-Path -Path $LogPath
    
    if (!$PathExists) {
        try {
            New-Item -Path $LogPath -ItemType "Directory"
        }
        catch {
            Write-Host "ERROR: Creating new directory: $LogPath"
        }
    }

    Return Test-Path -Path $LogPath
}