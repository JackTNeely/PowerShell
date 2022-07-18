# PowerShell
PowerShell, Remote Administration, Microsoft, Office, 365, Azure, Azure Active Directory, Bulk Scripting, Automation

- <b>Validate CNAME autodisocver record, modify client registry, and repair Outlook autodiscover based on type of Exchange services (using email address as input).</b>
<br />Note: This PowerShell script gets passed into cmd, allowing a seamless click-to-run user experience. 


<pre>
@echo off
PowerShell -Command "Get-Content '%~dpnx0' | Select-Object -Skip 3 | Out-String | Invoke-Expression"
goto :eof
Write-Host "`r`n`r`nWhat is your email address?" -ForegroundColor Cyan
$domain = (Read-Host).Split("@")[1]

Write-Host "`r`n`r`nFetching CNAME record for autodiscover.$domain . . ." -ForegroundColor Cyan
$d = Resolve-DnsName -Name autodiscover.$domain -Type CNAME -DnsOnly -QuickTimeout -ErrorAction SilentlyContinue
$d | ft Name,NameHost

$dS = ($d.NameHost -eq "autodiscover.outlook.com") -or ($d.NameHost -eq "adr.exghost.com")
if ($d.NameHost -eq $null ) { Write-Host "`r`n`r`nAutodiscover record not published. Please contact your administrator or try again." }
elseif (!$dS) { Write-Host "`r`n`r`nAutodiscover record not valid. Please contact your administrator or try again." }
else { Write-Host "`r`n`r`nAutodiscover record valid"

Write-Host "`r`n`r`nTesting connection for autodiscover.$domain . . ." -ForegroundColor Cyan
$p = Test-Connection autodiscover.$domain -Count 1 -Quiet

If (!$p) { Do { Write-Host "`r`nping autodiscover.$domain was NOT successful.`r`nTesting connection again . . ."
                                $p = Test-Connection autodiscover.$domain -Count 1 -Quiet
                                Write-Host "`r`n`r`n( Press Ctrl C to cancel )"
                                } until ($p) }
Write-Host "`r`nping autodiscover.$domain successful`r`n`r`n"

If ($dS -and $p -and $d.NameHost -eq "autodiscover.outlook.com") { Write-Host "No autodiscover issues detected.`r`nMicrosoft 365 Autodiscover Registry Edit is recommended." } 
elseif ($dS -and $p -and $d.NameHost -eq "adr.exghost.com") { Write-Host "No autodiscover issues detected.`r`nHosted Exchange Autodiscover Registry Edit is recommended." }
else { Write-Host "Something went wrong." } }

Sleep 30
</pre>

<br />
<br />
<br />

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
