#Need to connect to Exchange Online before running the script

#Pull the users in from the CSV file
$users = Import-Csv "users.csv"

#Create a folder for this minute
$folder = New-Item -ItemType Directory -Name (Get-Date -Format "dddd MM-dd-yyyy HH_mm")

#Change script location to folder
Set-Location $folder

#Loop through each user and generate two files - one for mailbox and one for archive
foreach ($user in $users)
    {
        #Narrow down to only the username - no column header
        $username = $user.name
        
        #Make a file with the user's mailbox statistics
        Get-MailboxStatistics -Identity $username | Format-List | Out-File "$username mailbox statistics.txt"

        #Make a file with the user's archive statistics
        Get-MailboxStatistics -Identity $username -Archive | Format-List | Out-File "$username archive statistics.txt"
    }

#Go back
Set-Location ..