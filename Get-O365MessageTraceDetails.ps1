Function Get-MessageTraceDetails {

    [CmdletBinding()] 

    param (
        [Parameter(Mandatory=$false)][string]$MessageID,
        [Parameter(Mandatory=$false)][string]$RecipientAddress,
        [Parameter(Mandatory=$false)][string]$SenderAddress,
        [Parameter(Mandatory=$false)][string]$StartDate,
        [Parameter(Mandatory=$false)][string]$EndDate
    )
    # You can use this cmdlet to search message data for the last 10 days. If you run this cmdlet without any parameters, only data from the last 48 hours is returned.
    if ($MessageID) {
        Return Get-MessageTrace -MessageID $MessageID | Get-MessageTraceDetail
    }
    else {
        Return Get-MessageTrace -StartDate $StartDate -EndDate $EndDate -SenderAddress $SenderAddress -RecipientAddress $RecipientAddress | Get-MessageTraceDetail
    }
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
