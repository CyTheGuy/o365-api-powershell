# https://support.office.com/en-us/article/Enable-mailbox-auditing-in-Office-365-aaca8987-5b62-458b-9882-c28476a66918#ID0EABAAA=Step-by-step_instructions
# Invalid audit operation specified. Supported audit operations for Mailbox Owner are None, Create, SoftDelete, HardDelete, Update, Move, MoveToDeletedItems and MailboxLogin.

Function Connect-ExchangeOnlineBasicAuth {

    try {
        $Cred = Get-Credential
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Cred -Authentication Basic -AllowRedirection
        Import-PSSession $Session
    }
    catch {
        Write-Host "Failed To Connect To Exchange Online With Basic Auth, Make Sure You Are Using Exchange Shell Or Running Powershell As Admin"
    }
}

Function Connect-ExchangeMFA {

  try {
      Connect-EXOPSSession
  }
  catch {
      Write-Host "Failed To Connect To Exchange Online Module"
      Write-Host "$($Error[0].Exception.Message)"
  }
}

function Update-AuditMailbox {

  param (
    [Parameter(Mandatory=$true)][array]$Users,
    [switch]$Enabled
  )

  foreach ($User in $Users.Alias) {
      if ($Enabled) {
        Write-Host "ADD: Mailbox Auditing for $User"
        Set-Mailbox -Identity $User -AuditEnabled $true -AuditOwner Create,HardDelete,MailboxLogin,Move,MoveToDeletedItems,SoftDelete,Update -AuditAdmin Copy,Create,FolderBind,HardDelete,MessageBind,Move,MoveToDeletedItems,SendAs,SendOnBehalf,SoftDelete,Update -AuditDelegate Create,FolderBind,HardDelete,Move,MoveToDeletedItems,SendAs,SendOnBehalf,SoftDelete,Update
      }
      else {
        Write-Host "REMOVE: Mailbox Auditing for $User"
        Set-Mailbox -Identity $User -AuditEnabled $false -AuditOwner None -AuditAdmin None -AuditDelegate None
      }
    }
}

try {
  Connect-ExchangeOnlineBasicAuth
}
catch {
  Connect-ExchangeMFA
}

$UsersWhoArentAudited = Get-Mailbox -ResultSize Unlimited | ?{$_.AuditEnabled -eq $False}

# Enable It
Update-AuditMailbox -Users $UsersWhoArentAudited -Enabled

# Disable It
#Update-PTAuditMailbox -User $Users