# Disable/Enable Remote Powershell In O365 - THIS DOES NOT IMPACT AZURE AD COMMANDS
Function Connect-ExchangeMFA {

    try {
        Connect-EXOPSSession
    }
    catch {
        Write-Host "Failed To Connect To Exchange Online Module"
        Write-Host "$($Error[0].Exception.Message)"
    }
}

function Update-RemotePowershell {

    param (
        [Parameter(Mandatory=$true)][string]$User,
        [switch]$Enabled
    )

    if ($Enabled) {
        Set-User $User -RemotePowerShellEnabled $True
        Write-Host "Set $User To RemotePowerShellEnabled = True"
    }
    else {
        Set-User $User -RemotePowerShellEnabled $False
        Write-Host "Set $User To RemotePowerShellEnabled = False"
    }
}

Connect-ExchangeMFA

$AllowedUsers = Get-ADGroupMember -Identity "people-who-can-have-powershell-sg" -Recursive

$UsersWithRemotePowershellEnabled = Get-User -Filter {RemotePowerShellEnabled -eq $True} -ResultSize Unlimited

# Disable It
foreach ($User in $UsersWithRemotePowershellEnabled.Name) {

    if ($AllowedUsers.Name -notcontains $User) {

        Write-Host "Disabling Remote Powershell For $User"
        Update-RemotePowershell $User
    }
}

foreach ($User in $AllowedUsers) {

    $UserInfo = Get-User -Identity $User.Name 

    if (!$UserInfo.RemotePowerShellEnabled) {

        Write-Host "Enabling Remote Powershell For $User"
        Update-RemotePowershell $User -Enabled 
    }
}