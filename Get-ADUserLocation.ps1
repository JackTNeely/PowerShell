# Run this module and pass individual users during script execution, hitting enter when done, or pass in users as arguments. 
 
# Example usage:
# cd <script location>
# .\Get-ADUserLocation.ps1 "Adele Vance", "Bob Ross", "Johnny Cage"

# Search by DisplayName, EmailAddress, SamAccountName, or DistinguishedName.
# You can also pass EmployeeID and Job Title as search parameters.

param(
    [Parameter(Mandatory=$true)] [array]$users
)

# Provide a comma separated list (array) of users as an argument during script creation, or here:
# $users = @(
#     "Abba Cadabra",
#     "Alice Bob",
#     "Azure Administrator"
# )

# Import the ActiveDirectory module if not already loaded.
Import-Module ActiveDirectory

# Generate a unique timestamp.
$date = (Get-Date -F O).Replace(":",".") + "Z"

# Define the CSV file path.
$path = "$env:HOME\ADUserLocationInformation-$date.csv"

# Create an array to hold the results.
$results = @()

# Create a counter variable for the progress bar.
$i = 1

# Loop through each user.
foreach ($user in $users) {
    # Store the username in a variable for error handling.
    $username = $user

    # Get all matching users and update the users variable with the retrieved users.
    $adUsers = Get-ADUser -Properties * -Filter "(DisplayName -like '*$username*') -or (EmailAddress -like '*$username*') `
                        -or (SamAccountName -like '*$username*') -or (DistinguishedName -like '*$username*') `
                        -or (EmployeeId -like '*$username*') -or (Title -like '*$username*')"
    
    # Create a second loop so all users are returned and not just the first match for a given input.
    foreach ($adUser in $adUsers) {
        Write-Progress -Activity "Retrieving users . . ." -Status "$i/$($adUsers.Count) users processed."

        # If the user is found, store their user properties in a PSObject.
        if ($adUser) {
            Write-Host "`n    Retrieved properties for $($adUser.DisplayName) . . . "

            $result = New-Object PSObject -Property ([ordered]@{
                "Display Name"     = $adUser.DisplayName
                "Email"            = $adUser.EmailAddress
                "Job Title"        = $adUser.Title
                "Department"       = $adUser.Department
                "Office"           = $adUser.Office
                "Telephone Number" = $adUser.TelephoneNumber
                "Company"          = $adUser.Company
                "Manager"          = if ($adUser.Manager) { $user.Manager } else { "" }
                "Description"      = $adUser.Description
                "Street Address"   = $adUser.StreetAddress
                "City"             = $adUser.City
                "State"            = $adUser.State
                "POBox"            = $adUser.POBox
                "Postal Code"      = $adUser.PostalCode
                "Country"          = $adUser.Country
            })

            # Add the result to the results array.
            $results += $result
        } else {
            Write-Host "`n    User Display Name '$username' was not found." -ForegroundColor Red
        }

        $i++
    }
}

# Export the results to a CSV file
$results | Export-Csv -Path $path -NoTypeInformation

Write-Host "`n    Exported AD user location information to" -ForegroundColor Green
Write-Host "    👉    $path" -ForegroundColor Yellow

Write-Host "`n`nPress Enter to open the CSV file report, or any other key to quit." -ForegroundColor Green
$input = Read-Host

if ($null -ne $input -or $input -eq $true -and $input -like "") { try { Start-Process $path } catch { Write-Error "An error occurred opening the CSV file." } } else { exit 0 }
