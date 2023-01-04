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
if ($d.NameHost -eq $null ) { Write-Host "`r`n`r`nAutodiscover record not published. Please contact your administrator or try again." -ForegroundColor Cyan }
elseif (!$dS) { Write-Host "`r`n`r`nAutodiscover record not valid. Please contact your administrator or try again." -ForegroundColor Cyan }
else { Write-Host "`r`n`r`nAutodiscover record valid"

Write-Host "`r`n`r`nTesting connection for autodiscover.$domain . . ." -ForegroundColor Cyan
$p = Test-Connection autodiscover.$domain -Count 1 -Quiet

If (!$p) { Do { Write-Host "`r`nping autodiscover.$domain was NOT successful.`r`nTesting connection again . . ."
                                $p = Test-Connection autodiscover.$domain -Count 1 -Quiet
                                Write-Host "`r`n`r`n( Press Ctrl C to cancel )"
                                } until ($p) }
Write-Host "`r`nping autodiscover.$domain successful`r`n`r`n"

if ($dS) {
            If ($dS -and $p -and $d.NameHost -eq "autodiscover.outlook.com") { Write-Host "No autodiscover issues detected.`r`nMicrosoft 365 Autodiscover Registry Edit is recommended." } 
            elseif ($dS -and $p -and $d.NameHost -eq "adr.exghost.com") { Write-Host "No autodiscover issues detected.`r`nHosted Exchange Autodiscover Registry Edit is recommended." }
            else { Write-Host "Something went wrong." } }

            Write-Host "`r`n`r`nBacking up registry . . ." -ForegroundColor Cyan

            Set-Location -Path HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\AutoDiscover

            $reg = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\AutoDiscover'
            $out = "$temp\backup.reg"

            reg export HKCU\SOFTWARE\Microsoft\Office\16.0\Outlook\AutoDiscover C:\Temp\backup.reg /y

            Write-Host "`r`n`r`n$reg backed up to $out" -ForegroundColor Cyan 
            Write-Host "`r`nPrevious registry can be restored by double clicking this file to run.`r`nThis will reverse the following registry changes:`r`n`r`n" 

            Set-ItemProperty -Path $reg -Name 'PreferLocalXML' -Value 0
            Set-ItemProperty -Path $reg -Name 'ExcludeHTTPRedirect' -Value 0
            Set-ItemProperty -Path $reg -Name 'PreferLocalXML' -Value 1
            Set-ItemProperty -Path $reg -Name 'ExcludeHttpsAutoDiscoverDomain' -Value 1
            Set-ItemProperty -Path $reg -Name 'ExcludeScpLookup' -Value 1
            Set-ItemProperty -Path $reg -Name 'ExcludeSrvLookup' -Value 1
            Set-ItemProperty -Path $reg -Name 'ExcludeSrvRecord' -Value 1


            if ($d.NameHost -eq "autodiscover.outlook.com") { 
                Set-ItemProperty -Path $reg -Name 'ExcludeExplicitO365Endpoint' -Value 0
                Write-Host "`r`n`r`nM365 Autodiscover Registry Edit successfully executed.`r`n`r`nProgram complete." -ForegroundColor Cyan }

            elseif ($d.NameHost -eq "adr.exghost.com") { 
                Set-ItemProperty -Path $reg -Name 'ExcludeExplicitO365Endpoint' -Value 1
                Write-Host "`r`n`r`nHosted Exchange Autodiscover Registry Edit successfully executed.`r`n`r`nProgram complete." -ForegroundColor Cyan }

            else { Write-Host "`r`n`r`nSomething went wrong. Please terminate this program." -ForegroundColor Cyan }
            }
Write-Host "`r`n`r`n`r`n`r`nPress enter to exit.`r`n`r`n`r`n" -ForegroundColor Magenta
Read-Host
</pre>

<br />
<br />
<br />

- <b>Sort HubSpot Leads List (CSV file) by Users@Domains using Microsoft 365 and Export to a New CSV</b>

<pre>
Write-Host "`r`n`r`n`r`n`r`nWhat is the full path for your file? `r`nE.g. C:\leads.csv" -ForegroundColor Magenta
$path = Read-Host -Prompt "Path"

$leads = Import-Csv -Path $path -Header "Record ID","First Name","Last Name","Email","Office Number","Primary Associated Company ID","Associated Company","Record ID - Company","Company name","Cell Number","City","Industry"
$count = $leads.count
$results = @()

Write-Host "`r`n`r`nProcessing . . . Do not exit." -ForegroundColor Magenta
Write-Host "This could take some time. Have a coffee. Maybe a few." -ForegroundColor Cyan
Start-Sleep -Seconds 3

$i = 0

foreach ($lead in $leads) {

    $domain = ($lead.Email).Split("@")[1]
    $autodiscover = Resolve-DnsName -Name autodiscover.$domain -Type CNAME -DnsOnly -QuickTimeout -ErrorAction SilentlyContinue
    $i++

    if ($autodiscover.NameHost -eq "autodiscover.outlook.com") {
        $results += $lead

        #Write-Host $domain -ForegroundColor Cyan
        #Write-Host $lead.Email -ForegroundColor Cyan
        #Write-Host $autodiscover.NameHost "`r`n"
        #Write-Host $results | ft

    } else {}

    Write-Host "$i out of $count"
}

$results | Export-Csv -Path C:\M365_Leads.csv -NoTypeInformation
Write-Host "`r`n`r`n`r`n`r`nScript complete!" -ForegroundColor Magenta
Write-Host "Results output to C:\M365_Leads.csv" -ForegroundColor Cyan
Write-Host "`r`nPress enter to exit." -ForegroundColor Magenta
</pre>

<br />
<br />

<div align="center">
<a href="https://lh3.googleusercontent.com/un87V2kkHHTlXhk6KwKgKygUCrtTzr4L-ikNGoCc5YZUIuBIKzQJ95o-70sTVKmcIQPCG6mBYBkHfLkrkeGp_Brli001wgi2wB-iNWhQa8yFHCb1e97a9eG-S8IKyWBm8Q_pszMfKPY=w2400?source=screenshot.guru" target="_blank"><img src="https://lh3.googleusercontent.com/un87V2kkHHTlXhk6KwKgKygUCrtTzr4L-ikNGoCc5YZUIuBIKzQJ95o-70sTVKmcIQPCG6mBYBkHfLkrkeGp_Brli001wgi2wB-iNWhQa8yFHCb1e97a9eG-S8IKyWBm8Q_pszMfKPY=w2400?source=screenshot.guru" />
</div>


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
