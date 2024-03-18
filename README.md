# Create365User 
## A PowerShell Script to create a new user to ActiveDirectory and Microsoft Office 365 Tenant

**1.** You will need to gather the Microsoft 365 Tenant ID.
<sub>
- https://portal.azure.com/ 
- Browse to Microsoft Entra ID > Properties.
- Scroll down to the Tenant ID section and you can find your tenant ID
</sub> 

**2.** You will need to create a Azure App that has access to the Exchange.
<sub>
</sub>

**3.** You will need to create an Azure User who has access to the Mailbox.
<sub>
</sub> 
MS Graph Scopes:  `
'Mail.ReadWrite'
'Mail.ReadBasic.All 
'User.ReadWrite.All' 
'Calendars.Read'
'Application.ReadWrite.All'
'Directory.ReadWrite.All'
'MailboxSettings.Read'
'Contacts.ReadWrite' 
'Directory.Read.All'
'User.Read.All'
'Organization.ReadWrite.All'
'Mail.Read'
'Calendars.ReadWrite' 
'LicenseAssignment.ReadWrite.All'
'Mail.Send'
'MailboxSettings.ReadWrite'
'Organization.Read.All'
'Contacts.Read'
'Mail.ReadBasic'
'Group.ReadWrite.All' `

**4.** You will need to create an Active Directory user which is a member of the "Domain Admins" Group.
<sub>
</sub> 
