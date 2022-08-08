#Get messages with failed status for further investigation

#Get today's date
$Today = Get-Date -Format "MM/dd/yy"

#Get the date of last week
$LastWeek = (Get-Date).AddDays(-7).ToShortDateString()

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

    #Get failed messages
Get-MessageTrace -StartDate $LastWeek -EndDate $Today | Where-Object { $_.Status -eq "Failed" } | Select-Object MessageTraceId,SenderAddress,RecipientAddress,Subject,Status | Export-Csv -Path ./failed_messages.csv
Import-Csv -Path ./failed_messages.csv

Write-Host "Failed messages exported to ./failed_messages.csv"