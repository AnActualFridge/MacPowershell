<#
.Synopsis
Hide the mailbox from the Global Address List.

.DESCRIPTION
This script will hide a given mailbox from the Global Address List.
.NOTES
Name: Hide-Mailbox
Author: R. Price 
Version: 1.0
DateCreated: June 2022
Purpose/Change: Hide Mailboxes

.EXAMPLE
Hide-Mailbox robert.price@contoso.com

Hides the address from the Global Address List.
#>
param (
    [Parameter(Mandatory = $true, HelpMessage = "The name of the mailbox to hide")]
    [string] $mailbox
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

#Get GAL status for mailbox
Write-Host "Current status for $mailbox is: " + (Get-Mailbox $mailbox | Select-Object HiddenFromAddressListsEnabled)
if ((Get-Mailbox $mailbox | Select-Object HiddenFromAddressListsEnabled) -eq "False") {
    Write-Host "Hiding $mailbox from GAL"
    Set-Mailbox -Identity $mailbox -HiddenFromAddressListsEnabled $true
} else {
    Write-Host "$mailbox is already hidden from Global address list.  No action taken."
}

#Ask whether to disconnect from Exchange Online
Write-Host "Disconnect from Exchange Online? (y/n)"
$Disconnect = Read-Host
if ($Disconnect -eq "y") {
    Disconnect-ExchangeOnline
}
else {
    Write-Host "Not disconnecting from Exchange Online.  Please use Disconnect-ExchangeOnline to disconnect when finished."
}
