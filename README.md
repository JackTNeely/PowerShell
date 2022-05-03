# PowerShell
PowerShell, Remote Administration, Microsoft, Office, 365, Azure, Azure Active Directory, Bulk Scripting, Automation

- <b>Get All Users by Mailbox</b>
- <b>Get User Principal Names</b>
- <b>Get Creation Dates</b>
- <b>Get Most-Recent-Change Dates</b>
- <b>Write to Host</b>
  
  
$objUsers = Get-Mailbox -ResultSize Unlimited | Select UserPrincipalName 
 
Foreach ($objUser in $objUsers) 
    {     
        $objUserMailbox = Get-User -Identity $($objUser.UserPrincipalName) | Select Identity, WhenCreated, WhenChanged
 
$strUserPrincipalName = $objUser.UserPrincipalName
$strWhenCreated = $objUserMailbox.whenCreated
$strWhenChanged = $objUserMailbox.whenChanged
write-host "$strUserPrincipalName : $strWhenCreated : $strWhenChanged"
    }

<div align="center">
  <h2>Much More to Come!</h2>
</div>
