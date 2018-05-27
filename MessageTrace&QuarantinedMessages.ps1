Function Connect-ExchangeMFA {

    try {
        Connect-EXOPSSession
    }
    catch {
        Write-Host "Failed To Connect To Exchange Online Module"
        Write-Host "$($Error[0].Exception.Message)"
    }
}
Function Get-MessageTraceDetails {

    [CmdletBinding()] 

    param (
        [Parameter(Mandatory=$true)][string]$MessageID
    )
    
    Return Get-MessageTrace -MessageID $MessageID | Get-MessageTraceDetail
}

Function Get-MessageQuarantineDetails {

    [CmdletBinding()] 

    param (
        [Parameter(Mandatory=$false)][string]$MessageID,
        [Parameter(Mandatory=$false)][string]$RecipientAddress,
        [Parameter(Mandatory=$false)][string]$SenderAddress,
        [Parameter(Mandatory=$false)][string]$Subject,
        [Parameter(Mandatory=$false)][string]$MyItems,
        [Parameter(Mandatory=$false)][string]$StartReceivedDate,
        [Parameter(Mandatory=$false)][string]$EndReceivedDate
    )
    
    Return Get-QuarantineMessage -MessageID $MessageID | FL
}