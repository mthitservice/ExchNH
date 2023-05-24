

Import-Module ExchangeOnlineManagement



Connect-ExchangeOnline -UserPrincipalName Michael.Lindner@mth-it-service.com
$t=Get-Mailbox
$t | ft

#Get-ExoMailboxStatistics -Identity Michael.Lindner |select DisplayName,TotalItemSize | Export-Csv -Path size.csv

#Get-Mailbox -ResultSize Unlimited | Get-MailboxStatistics | select DisplayName,TotalItemSize

####Get-Mailbox -ResultSize Unlimited -Archive
#Get-Mailbox -ResultSize Unlimited -Archive |  Get-MailboxStatistics| Select DisplayName,TotalItemSize

Get-Mailbox -Identity Michael.Lindner | Select *quota*

Get-EXOMailbox  -ResultSize Unlimited -RecipientTypeDetails SharedMailbox

#Get-Mailbox | foreach {
  ###  (Get-MailboxPermission -Identity $_.userprincipalname | where{ ($_.AccessRights -contains "FullAccess") -and ($_.IsInherited -eq $false) -and -not ($_.User -match "NT AUTHORITY") }) | select Identity,AccessRights,User}


 #   Get-Mailbox -ResultSize Unlimited |Foreach{
   #     Get-MailboxStatistics -Identity $_.UserPrincipalName | Select DisplayName,LastLogonTime,LastUserActionTime}

        Get-mailbox -ResultSize Unlimited| where {$_.ForwardingAddress -ne $Null} | select DisplayName,ForwardingAddress
        Get-MailboxFolder -Identity "Michael.Lindner@mth-it-service.com" -GetChildren | ft

        Get-MailboxFolderPermission -Identity "Michael.Lindner@mth-it-service.com:\LinkedIn"