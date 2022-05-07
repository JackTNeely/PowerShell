# PowerShell
PowerShell, Remote Administration, Microsoft, Office, 365, Azure, Azure Active Directory, Bulk Scripting, Automation

- <b>Bulk Remove Automapping for All Users</b>
  
<pre>  
$objUsers = Get-Mailbox -Resultsize Unlimited
Foreach ($objUser in $objUsers){
        $RemoveAutoMapping = Get-MailboxPermission -Identity $($objUser.Alias) | Where 
            {$.AccessRights -eq "FullAccess"} 
        $RemoveAutoMapping | Remove-MailboxPermission
        $RemoveAutoMapping | Foreach 
            {Add-MailboxPermission -Identity $.Identity -User $.User -AccessRights:FullAccess -AutoMapping $False}
    }
</pre>

<br/>
<br/>

<div align="center">
  <h2>Much More to Come!</h2>
</div>
