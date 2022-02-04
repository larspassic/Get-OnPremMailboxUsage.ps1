#Run this script on an Exchange Server in the Exchange Management Shell (EMS)

#Pull the users in from the CSV file
$users = Import-Csv "users.csv"

#Create a folder for this minute
$date = Get-Date -Format "dddd MM-dd-yyyy HH_mm"

#Make an empty array to store all the results
$array = @()

#Function inspired by Paul Cunningham
#https://github.com/cunninghamp/Powershell-Exchange/blob/master/Set-MailboxQuota/Set-MailboxQuota.ps1#L97
Function Convert-QuotaStringToGB() {

    Param([string]$CurrentQuota)

    [string]$CurrentQuota = ($CurrentQuota.Split("("))[1]
    [string]$CurrentQuota = ($CurrentQuota.Split(" bytes)"))[0]
    $CurrentQuota = $CurrentQuota.Replace(",","")
    
    #Format the string in to an int
    [int]$CurrentQuotaInBytes = "{0:F0}" -f ($CurrentQuota)
    
    #Take int of bytes and round it into GB as a decimal ie. 0.02
    [decimal]$CurrentQuotaInGB = [math]::round(($CurrentQuotaInBytes)/1GB, 2)
    
    return $CurrentQuotaInGB
}

#Loop through each user and make a CSV of the primary and archive mailbox usage details
foreach ($user in $users)
    {
        #Collect the user's primary mailbox statistics
        $statistics = Get-MailboxStatistics -Identity $user.UserPrincipalName | Select-Object -Property "UserPrincipalName", @{n="PrimaryNormalItemSizeGB";e={$_.TotalItemSize}},"NewProhibitSendReceiveQuotaGB", @{n="PrimaryRecoverableItemSizeGB";e={$_.TotalDeletedItemSize}},"NewRecoverableItemsQuotaGB", "ArchiveDisplayName", @{n="ArchiveNormalItemSizeGB";e={$_.ArchiveTotalItemSize}},"NewArchiveQuotaGB",@{n="ArchiveRecoverableItemSizeGB";e={$_.ArchiveTotalDeletedItemSize}},"NewArchiveRecoverableItemsQuotaGB"
        
        #Set the primary mailbox statistics
        $statistics.UserPrincipalName = $user.UserPrincipalName
        $statistics.PrimaryNormalItemSizeGB = Convert-QuotaStringToGB -CurrentQuota $statistics.PrimaryNormalItemSizeGB
        $statistics.PrimaryRecoverableItemSizeGB = Convert-QuotaStringToGB -CurrentQuota $statistics.PrimaryRecoverableItemSizeGB

        #Collect the user's archive statistics
        $archivestatistics = Get-MailboxStatistics -Identity $user.UserPrincipalName -Archive 
        
        #Assign the archive columns to their values
        $statistics.ArchiveDisplayName = $archivestatistics.DisplayName
        $statistics.ArchiveRecoverableItemSizeGB = Convert-QuotaStringToGB -CurrentQuota $archivestatistics.TotalDeletedItemSize
        $statistics.ArchiveNormalItemSizeGB = Convert-QuotaStringToGB -CurrentQuota $archivestatistics.TotalItemSize

        #Add the current user to the array
        $array += $statistics
    }

#Export the array as a complete CSV
$array | Export-Csv "OnPrem Primary and Archive Statistics $date.csv" -NoTypeInformation

#End of script