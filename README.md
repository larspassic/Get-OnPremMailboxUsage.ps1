# Get-OnPremMailboxUsage.ps1

## Introduction
This script uses the **Get-MailboxStatistics** Exchange Management Shell cmdlet to gather  vital mailbox usage statistics about a batch of users.

| Mailbox type | Value from Get-MailboxStatistics | Column name in output file |
| - | - | - |
| Primary mailbox | TotalItemSize | PrimaryNormalItemSizeGB
| Primary mailbox | TotalDeletedItemSize | PrimaryRecoverableItemSizeGB
| Archive mailbox | TotalItemSize | ArchiveNormalItemSizeGB
| Archive mailbox | TotalDeletedItemSize | ArchiveRecoverableItemSizeGB

This script converts the string values of mailbox sizes produced by the **Get-MailboxStatistics** cmdlet, into decimal values (in gigabytes), so that the resulting output file can be quickly sorted. 

The script also also creates empty columns so that you can easily plan for quota adjustments, if needed.

## Instructions
1) Populate **users.csv** with UserPrincipalNames (UPNs) of the current batch of users.
1) Run the script from the Exchange Management Shell (EMS) on an Exchange server in your environment.
1) The script will create a csv file named **"OnPrem Mailbox Usage MM-DD-YYYY HH_MM.csv"** with the necessary data.

Please use the latest release from the "Releases" section on the right-hand side of this page.

## Credits
The function **Convert-QuotaStringToGB()** in this script was inspired by Paul Cunningham ([cunninghamp](https://github.com/cunninghamp)) from his function **[Convert-QuotaStringToKB()]((https://github.com/cunninghamp/Powershell-Exchange/blob/master/Set-MailboxQuota/Set-MailboxQuota.ps1#L97))** from his script **Set-MailboxQuota.ps1**.
