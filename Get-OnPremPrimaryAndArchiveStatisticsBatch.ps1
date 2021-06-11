#Run this script on an Exchange Server in the Exchange Management Shell (EMS)

#Pull the users in from the CSV file
$users = Import-Csv "users.csv"

#Create a folder for this minute
$date = Get-Date -Format "dddd MM-dd-yyyy HH_mm"

#Make an empty array to store all the results
$array = @()

#Loop through each user and make a CSV of the primary and archive mailbox usage details
foreach ($user in $users)
    {
        #Collect the user's primary mailbox statistics and create 3 columns for archive statistics
        $statistics = Get-MailboxStatistics -Identity $user.upn | Select-Object -Property DisplayName, @{n="PrimaryTotalDeletedItemSize";e={$_.TotalDeletedItemSize}}, @{n="PrimaryTotalItemSize";e={$_.TotalItemSize}}, "ArchiveDisplayName", "ArchiveTotalDeletedItemSize", "ArchiveTotalItemSize"
        
        #Collect the user's archive statistics
        $archivestatistics = Get-MailboxStatistics -Identity $user.upn -Archive 
        
        #Assign the archive columns to their values
        $statistics.ArchiveDisplayName = $archivestatistics.DisplayName
        $statistics.ArchiveTotalDeletedItemSize = $archivestatistics.TotalDeletedItemSize
        $statistics.ArchiveTotalItemSize = $archivestatistics.TotalItemSize

        #Add the current user to the array
        $array += $statistics
    }

#Export the array as a complete CSV
$array | Export-Csv "OnPrem Primary and Archive Statistics $date.csv" -NoTypeInformation

#End of script