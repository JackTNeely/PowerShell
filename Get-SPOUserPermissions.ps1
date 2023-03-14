param (
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,

    [Parameter(Mandatory=$true)]
    [string]$UserEmail,

    [Parameter(Mandatory=$true)]
    [string]$OutputCSVFile
)

function Connect-SPO {
    param (
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl
    )

    try {
        $credential = Get-Credential
        Connect-SPOService -Url $SiteUrl -Credential $credential
    } catch {
        Write-Error "Failed to connect to SharePoint Online. Please ensure your credentials are correct and try again."
        exit 1
    }
}

function Get-UserPermissions {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserEmail
    )

    try {
        $web = Get-SPOSite -Identity $SiteUrl -Detailed
        $user = Get-SPOUser -Site $SiteUrl -LoginName $UserEmail

        $permissionLevels = @()

        foreach ($roleAssignment in $web.RoleAssignments) {
            if ($roleAssignment.Member.LoginName -eq $user.LoginName) {
                $permissionLevels += $roleAssignment.RoleDefinitionBindings.Title
            }
        }

        return $permissionLevels
    } catch {
        Write-Error "Failed to retrieve user permissions. Please ensure the user email and site URL are correct."
        exit 1
    }
}

function Export-PermissionsToCSV {
    param (
        [Parameter(Mandatory=$true)]
        [Array]$Permissions,

        [Parameter(Mandatory=$true)]
        [string]$OutputCSVFile
    )

    try {
        $permissionsTable = $Permissions | Select-Object @{Name='Permission'; Expression={$_}}

        $permissionsTable | Export-Csv -Path $OutputCSVFile -NoTypeInformation
        Write-Host "Permissions exported to $OutputCSVFile" -ForegroundColor Green
    } catch {
        Write-Error "Failed to export permissions to CSV file."
        exit 1
    }
}

Connect-SPO -SiteUrl $SiteUrl
$permissions = Get-UserPermissions -UserEmail $UserEmail
Export-PermissionsToCSV -Permissions $permissions -OutputCSVFile $OutputCSVFile
