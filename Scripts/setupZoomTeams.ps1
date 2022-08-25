#Quick & dirty script that sets up 'Resources' in 365
#so zoom devices can accept teams meetings.
#Settings also need to be changed via Zoom Admin to 
#allow direct guest connections

#Get meeting room email address as parameter
Param (
    [Parameter(Mandatory = $true)]
    [string] $meetingRoomEmailAddress
)

#Check current settings for the meeting room
Get-CalendarProcessing -Identity $meetingRoomEmailAddress | Format-List Identity,DeleteComments,DeleteSubject,AddOrganizerToSubject,RemovePrivateProperty,DeleteAttachments     

#Confirm from input whether to continue
Write-Host "Update current settings? (Y/N)"
$continue = Read-Host
if ($continue -eq "Y") {
    #Change settings
    Write-Host "Changing settings"
    Set-CalendarProcessing -Identity $meetingRoomEmailAddress -AddOrganizerToSubject $false -OrganizerInfo $true -DeleteComments $false -DeleteSubject $false -RemovePrivateProperty $false
} else {
    #Continue with current settings
    Write-Host "Continuing with current settings"
}

#Print current settings
Get-CalendarProcessing -Identity $meetingRoomEmailAddress | Format-List Identity,DeleteComments,DeleteSubject,AddOrganizerToSubject,RemovePrivateProperty,DeleteAttachments
