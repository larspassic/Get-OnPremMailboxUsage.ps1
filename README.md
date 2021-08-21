# Get-OnPremPrimaryAndArchiveStatisticsBatch.ps1

## This script uses the **Get-MailboxStatistics** Exchange Management Shell cmdlet to gather six vital statistics about a batch of users.

1) Please populate **users.csv** with UserPrincipalNames (UPNs) of the current batch of users.
1) Please run the script from the Exchange Management Shell (EMS) on an Exchange server in your environment.
1) The script will create a csv file named **"OnPrem Primary and Archive Statistics Friday MM-DD-YYYY HH_MM.csv"** with the necessary data.



Please use the latest release from the "Releases" section on the right-hand side of this page.
