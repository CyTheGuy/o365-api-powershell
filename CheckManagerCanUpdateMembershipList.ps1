# Pass In SAMAccount Name And Get Managers
$Groups = Get-ADGroup -Filter * -SearchBase "OU=SelfService-DistGroups,DC=YOJOE,DC=Local" -Properties MailNickName,msExchCoManagedByLink,ManagedBy | Select MailNickName,msExchCoManagedByLink,ManagedBy | Sort MailNickName

Foreach ($Group in $Groups) {

    $Owners = @()

    if (!$Group.ManagedBy) {

        Write-Host "$($Group.MailNickName) Doesnt Have Anything In ManagedBy"
    }
    if (!$Group.msExchCoManagedByLink -and !$Group.ManagedBy) {

        Write-Host "$($Group.MailNickName) Doesnt Have Anything In msExchCoManagedByLink or ManagedBy Skipping Them"
        Continue
    }

    foreach ($Owner in $Group.msExchCoManagedByLink) {

        $Owners += $Owner
    }

    # Add Them To An Array
    $Owners += $Group.ManagedBy 

    # Need To Get Name With Get-DistributionGroup
    $GroupName = Get-DistributionGroup -Identity $Group.MailNickName | Select Name

    # Using The Name From EAC Add Perms
    Write-Host "Checking Box For $($Group.MailNickName)"
    $Owners | %{Add-ADPermission -Identity $GroupName.Name -User $_ -AccessRights WriteProperty -Properties “Member” -WarningAction 'silentlyContinue' | Out-Null} 
}
