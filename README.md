# Create a new 365 and Active Directory User.
A PowerShell Script to create a new user in an ActiveDirectory Domain or a Microsoft Office 365 Tenant.

![Screenshot of Loading Create New 365 User.](./Loading-Create365User.png)
![Screenshot of a Create New 365 User.](./CreateNewUser.png)

<sub> References: 
- [Microsoft.Online.SharePoint.PowerShell](https://learn.microsoft.com/en-us/powershell/module/sharepoint-online/connect-sposervice?view=sharepoint-ps)
- https://learn.microsoft.com/en-us/power-apps/developer/data-platform/walkthrough-register-app-azure-active-directory
- https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.authentication/invoke-mggraphrequest?view=graph-powershell-1.0
</sub> 

# Instructions

**1.** You will need to gather the Microsoft 365 Tenant ID.
<sub>
- https://portal.azure.com/ 
- Browse to Microsoft Entra ID > Properties.
- Scroll down to the Tenant ID section and you can find your tenant ID
</sub> 

**2.** You will need to create a 365 Entra App that has access to the Exchange.
<sub>Microsoft Graph		User.Read						Sign in and read user profile					Delegated		Admin consent	An administrator</sub>
<sub>Microsoft Graph		Mail.ReadWrite					Read and write mail in all mailboxes			Application		Admin consent	An administrator</sub>
<sub>Microsoft Graph		User.ReadWrite.All				Read and write all users' full profiles			Application		Admin consent	An administrator</sub>
Microsoft Graph		Calendars.Read					Read calendars in all mailboxes					Application		Admin consent	An administrator
Microsoft Graph		Mail.ReadBasic.All				Read basic mail in all mailboxes				Application		Admin consent	An administrator
Microsoft Graph 	Application.ReadWrite.All		Read and write all applications					Application		Admin consent	An administrator
Microsoft Graph		Directory.ReadWrite.All			Read and write directory data					Application		Admin consent	An administrator
Microsoft Graph		MailboxSettings.Read			Read all user mailbox settings					Application		Admin consent	An administrator
Microsoft Graph		Sites.ReadWrite.All				Read and write items in all site collections	Application		Admin consent	An administrator
Microsoft Graph		Contacts.ReadWrite				Read and write contacts in all mailboxes		Application		Admin consent	An administrator
Microsoft Graph		Group.ReadWrite.All				Read and write all groups						Application		Admin consent	An administrator
Microsoft Graph		Directory.Read.All				Read directory data								Application		Admin consent	An administrator
Microsoft Graph		User.Read.All					Read all users' full profiles					Application		Admin consent	An administrator
Microsoft Graph		Organization.ReadWrite.All		Read and write organization information			Application		Admin consent	An administrator
Microsoft Graph		Mail.Read						Read mail in all mailboxes						Application		Admin consent	An administrator
Microsoft Graph		Calendars.ReadWrite				Read and write calendars in all mailboxes		Application		Admin consent	An administrator
Microsoft Graph		LicenseAssignment.ReadWrite.All	Manage all license assignments					Application		Admin consent	An administrator
Microsoft Graph		Mail.Send						Send mail as any user							Application		Admin consent	An administrator
Microsoft Graph		MailboxSettings.ReadWrite		Read and write all user mailbox settings		Application		Admin consent	An administrator
Microsoft Graph		Organization.Read.All			Read organization information					Application		Admin consent	An administrator
Microsoft Graph		GroupMember.ReadWrite.All		Read and write all group memberships			Application		Admin consent	An administrator
Microsoft Graph		Contacts.Read					Read contacts in all mailboxes					Application		Admin consent	An administrator
Microsoft Graph		Mail.ReadBasic					Read basic mail in all mailboxes				Application		Admin consent	An administrator

Office 365 Exchange Online	Office 365 Exchange Online	full_access_as_app				Use Exchange Web Services with full access to all mailboxes	Application	Admin consent	An administrator
Office 365 Exchange Online	Exchange.ManageAsApp		Manage Exchange As Application				Application		Admin consent	An administrator
</sub>

**3.** You will need to create an Entra User who has full access to the Tenants Sharepoint
<sub> [Microsoft.Online.SharePoint.PowerShell](https://learn.microsoft.com/en-us/powershell/module/sharepoint-online/connect-sposervice?view=sharepoint-ps) </sub> 
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

**4.** You will need to create an SelfSigned Certificate wich is adde to the 365 App
```powershell
$cert = New-SelfSignedCertificate -DnsName "USE: Create365User.ini -> ConnectSPOServiceUser" -CertStoreLocation cert:\LocalMachine\My -Type SSLServerAuthentication -NotAfter 2024-01-01 -NotBefore 2029-01-01
$pwd = ConvertTo-SecureString -String "USE: Create365User.ini -> ActiveDirectoryPassword" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath "USE: Create365User.ini -> ConnectSPOServiceUser.pfx" -Password $pwd
```

**5.** You will need to create an Active Directory user which is a member of the "Domain Admins" Group.
<sub>
</sub> 

**6.** Gather all the Tenants Licence IDs.
<sub>
</sub> 

**7.** Update the Create365User.ini file using all the details previously gathered.
<sub>
</sub> 
