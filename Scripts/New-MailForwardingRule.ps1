<#
.Synopsis
Create a new mail-forwarding rule.

.DESCRIPTION
This script will create a new mail-forwarding rule.
Note that the date relies on the format in this computer's regional settings.
Currently uses AEST (+10).

.NOTES
Name: New-MailForwardingRule
Author: R. Price
Version: 1.0
DateCreated: July 2022
Purpose/Change: Mail Forwarding

.EXAMPLE
New-MailForwardingRule -From rob@contoso.com.au -To eliot@contoso.com

Forwards all mail from first address to second address indefinitely.

New-MailForwardingRule -From rob@contoso.com.au -To webster@contoso.com.au -Start "09/01/2018 5:00 PM" -End "15/01/2018 8:00 AM"
Forwards all mail from first address to second address from 9/1/2018 5:00 PM to 15/1/2018 8:00 AM AEST.
#>
param (
    [Parameter(Mandatory = $true, HelpMessage = "The name of the mailbox to forward from")]
    [string] $From,

    [Parameter(Mandatory = $true, HelpMessage = "The name of the mailbox to forward to")]
    [string] $To,

    [Parameter(Mandatory = $false, HelpMessage = "The start date and time of the rule")]
    [string] $Start,

    [Parameter(Mandatory = $false, HelpMessage = "The end date and time of the rule")]
    [string] $End,

    [Parameter(Mandatory = $false, HelpMessage = "Messages will be kept if this flag is used")]
    [bool] $KeepMessages = $false
)

#Are we connected to Exchange Online?  If not, connect
Write-Host "Testing connection to Exchange Online"
try {
    Get-Mailbox test@example.com
}
catch [System.Management.Automation.CommandNotFoundException]{
    Write-Warning "Connecting to Exchange Online... please authenticate"
    Connect-ExchangeOnline 
}
catch [Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException]{
    Write-Warning "Already connected to Exchange Online.  Please use Disconnect-ExchangeOnline to disconnect when finished."
}
Write-Host "Connected to Exchange Online, getting messages..."

#Check the From address is valid
$FromAddress = Get-CASMailbox $From
If ($null -eq $FromAddress) {
    Write-Host "From address is not valid"
    exit 1
} else {
    Write-Host "From address is valid"
}

#Check the To address is valid
$ToAddress = Get-CASMailbox $To
If ($null -eq $ToAddress) {
    Write-Host "To address is not valid"
    exit 1
} else {
    Write-Host "To address is valid"
}

#If start time exists, check it is valid
If ($Start -ne $null) {
    $StartDateTime = Get-Date -Date $Start
    If ($null -eq $StartDateTime) {
        Write-Host "Start time is not valid; try using the format: dd/mm/yyyy hh:mm AM/PM"
        exit 1
    }
}

#If end time exists, check it is valid
If ($End -ne $null) {
    $EndDateTime = Get-Date -Date $End
    If ($null -eq $EndDateTime) {
        Write-Host "End time is not valid; try using the format: dd/mm/yyyy hh:mm AM/PM"
        exit 1
    }
}

#Adjust times for AEST
$StartDateTime = $StartDateTime.AddHours(-10)
$EndDateTime = $EndDateTime.AddHours(-10)

#Create the inbox rule for the From address
Write-Host "Attempting to create inbox rule for $From"
$InboxRule = New-InboxRule -Mailbox $From -Name "Forwarding to $To Rule" -RedirectTo $To -ReceivedAfterDate $StartDateTime -ReceivedBeforeDate $EndDateTime
Write-Host "Inbox rule created: " $InboxRule

#Ask whether to disconnect from Exchange Online
Write-Host "Disconnect from Exchange Online? (y/n)"
$Disconnect = Read-Host
if ($Disconnect -eq "y") {
    Disconnect-ExchangeOnline
}
else {
    Write-Host "Not disconnecting from Exchange Online"
}
