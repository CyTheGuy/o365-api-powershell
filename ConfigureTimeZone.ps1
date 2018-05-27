# This script updates the timezone for a user in O365 look at the quip doc linked below for more info on parameters
# Codes: https://www.cogmotive.com/blog/office-365-tips/set-language-and-time-zone-for-all-owa-users-in-office-365

Function Connect-ExchangeMFA {

    try {
        Connect-EXOPSSession
    }
    catch {
        Write-Host "Failed To Connect To Exchange Online Module"
        Write-Host "$($Error[0].Exception.Message)"
    }
}
Function Get-TimeZone {

    param (
        [Parameter(Mandatory=$true)][string]$Email
    )

    try {
        Get-MailboxRegionalConfiguration -Identity $Email
    }
    catch {
        Write-Host "Failed To Get Time Zone Config For $Email Make Sure You Passed An Email Address That Is In An O365 Tenant"
        Write-Host "$($Error[0].Exception.Message)"
    }
}

Function Update-TimeZone {

    param (
        [Parameter(Mandatory=$true)][string]$Email,
        [string]$TimeZone = "Pacific Standard Time",
        [string]$TimeFormat = "h:mm tt",
        [string]$DateFormat = "M/d/yyyy"
    )

    try {
        Set-MailboxRegionalConfiguration -Identity $Email -TimeZone $TimeZone -TimeFormat $TimeFormat -DateFormat $DateFormat -Confirm:$False
    }
    catch {

        Write-Host "Failed To Set Time Zone Config For $Email Make Sure You Passed An Email Address In An O365 Tenant And A Correct TimeZone"
        Write-Host "Use https://www.cogmotive.com/blog/office-365-tips/set-language-and-time-zone-for-all-owa-users-in-office-365 As A Reference For Timezone Parameters"
        Write-Host "$($Error[0].Exception.Message)"
    }
}