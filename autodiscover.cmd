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
