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