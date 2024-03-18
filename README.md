# Create365User 
## a PowerShell Script to add a new user to ActiveDirectory and 365

You will need to gather the Micirtosoft 365 Tenant ID

You will need to create a Azure App that has access to the Exchance

You will need to create an Azure User who has access to the Mailbox 
MS Graph Scopes: 'Mail.ReadWrite','Mail.ReadBasic.All, 'User.ReadWrite.All','Calendars.Read',','Application.ReadWrite.All','Directory.ReadWrite.All','MailboxSettings.Read','Contacts.ReadWrite','Directory.Read.All','User.Read.All','Organization.ReadWrite.All','Mail.Read','Calendars.ReadWrite','LicenseAssignment.ReadWrite.All','Mail.Send','MailboxSettings.ReadWrite','Organization.Read.All','Contacts.Read','Mail.ReadBasic','Group.ReadWrite.All'

You will need to create an Active Directory user which is a member of the "Domain Admin" Group.
