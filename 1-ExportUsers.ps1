<#
.SYNOPSIS
    Exports licensed users from the source M365 tenant to a CSV file.

.DESCRIPTION
    This script connects to the source M365 tenant using the Microsoft Graph PowerShell SDK,
    retrieves all users, filters for those with assigned licenses, and exports their details
    along with their license SKU IDs to a CSV file.

    Requires (minimum Graph scopes): User.Read.All, Directory.Read.All, Organization.Read.All (delegated, interactive sign-in; Organization.Read.All is needed by the tenant check).

.PARAMETER ConfigPath
    Path to the migration configuration file. Defaults to '.\MigrationConfig.psd1'.

.EXAMPLE
    .\1-ExportUsers.ps1 -ConfigPath '.\MigrationConfig.psd1'
#>

[CmdletBinding()]
param (
    [string]$ConfigPath = '.\MigrationConfig.psd1'
)

. "$PSScriptRoot\MigrationHelpers.ps1"

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
$env:AZURE_IDENTITY_DISABLE_WAM = 'true'

# Define paths relative to the script directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutputDir = Join-Path $ScriptDir "Output"
$LogDir = Join-Path $ScriptDir "Logs"

# Create directories if they don't exist
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }

$LogFile = Join-Path $LogDir "ExportUsers_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = 'INFO'
    )
    $LogMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogMessage
    
    switch ($Level) {
        'INFO' { Write-Host $LogMessage -ForegroundColor Green }
        'WARN' { Write-Host $LogMessage -ForegroundColor Yellow }
        'ERROR' { Write-Host $LogMessage -ForegroundColor Red }
        'DEBUG' { Write-Verbose $LogMessage }
        Default { Write-Host $LogMessage }
    }
}

try {
    Write-Log "Starting User Export Process..."
    
    # Load configuration
    if (Test-Path $ConfigPath) {
        Write-Log "Loading configuration from $ConfigPath"
        $Config = Import-PowerShellDataFile -Path $ConfigPath
    } else {
        Write-Log "Configuration file not found at $ConfigPath. Continuing without it." -Level WARN
    }

    Write-Log "Connecting to Source Tenant via Microsoft Graph..."
    Write-Log "A browser window will open. Sign in with your SOURCE tenant admin account." -Level WARN
    $sourceTenant = Get-RequiredConfigValue -Value $Config.SourceTenantDomain -Name 'SourceTenantDomain'
    $connectParams = @{ Scopes = @('User.Read.All','Directory.Read.All','Organization.Read.All'); NoWelcome = $true; TenantId = $sourceTenant }
    Connect-MgGraph @connectParams
    Assert-Tenant -ExpectedTenantDomain $sourceTenant
    Write-Log "Connected successfully. Tenant verified against SourceTenantDomain."

    # Define user properties to retrieve
    $UserProperties = @(
        "Id", "DisplayName", "UserPrincipalName", "Mail", "MailNickname",
        "GivenName", "Surname", "JobTitle", "Department", "OfficeLocation",
        "UsageLocation", "AccountEnabled", "CompanyName", "City", "Country",
        "PostalCode", "State", "StreetAddress", "MobilePhone", "BusinessPhones"
    )

    Write-Log "Retrieving all users from the tenant..."
    $AllUsers = Get-MgUser -All -Property $UserProperties
    $TotalUsers = $AllUsers.Count
    Write-Log "Found $TotalUsers total users in the tenant."

    $ExportedUsers = [System.Collections.Generic.List[object]]::new()
    $LicenseBreakdown = @{}

    Write-Log "Filtering users and retrieving license details..."
    $Counter = 0
    foreach ($User in $AllUsers) {
        $Counter++
        Write-Progress -Activity "Processing Users" -Status "User $Counter of $TotalUsers" -PercentComplete (($Counter / $TotalUsers) * 100)

        # Get license details for the user (S18: retry with backoff on 429/transient errors)
        $userId = $User.Id
        $LicenseDetails = Invoke-WithRetry -ScriptBlock { Get-MgUserLicenseDetail -UserId $userId }

        # S18: light throttle so a large tenant does not trip Graph throttling (429)
        if ($Counter % 10 -eq 0) {
            Start-Sleep -Milliseconds 500
        }

        if ($LicenseDetails) {
            $SkuIds = $LicenseDetails.SkuId -join ';'
            
            # Count licenses for breakdown
            foreach ($SkuId in $LicenseDetails.SkuId) {
                if ($LicenseBreakdown.ContainsKey($SkuId)) {
                    $LicenseBreakdown[$SkuId]++
                } else {
                    $LicenseBreakdown[$SkuId] = 1
                }
            }

            # Prepare the user object for export
            $ExportUser = [PSCustomObject]@{
                Id                = $User.Id
                DisplayName       = $User.DisplayName
                UserPrincipalName = $User.UserPrincipalName
                MailNickname      = $User.MailNickname
                GivenName         = $User.GivenName
                Surname           = $User.Surname
                Mail              = $User.Mail
                JobTitle          = $User.JobTitle
                Department        = $User.Department
                OfficeLocation    = $User.OfficeLocation
                UsageLocation     = $User.UsageLocation
                AccountEnabled    = $User.AccountEnabled
                CompanyName       = $User.CompanyName
                City              = $User.City
                Country           = $User.Country
                PostalCode        = $User.PostalCode
                State             = $User.State
                StreetAddress     = $User.StreetAddress
                MobilePhone       = $User.MobilePhone
                BusinessPhones    = if ($User.BusinessPhones) { $User.BusinessPhones -join ';' } else { $null }
                LicenseSkuIds     = $SkuIds
            }

            $ExportedUsers.Add($ExportUser)
        }
    }
    
    $LicensedUsersCount = $ExportedUsers.Count
    Write-Log "Found $LicensedUsersCount licensed users to export."

    $CsvPath = Join-Path $OutputDir "ExportedUsers.csv"
    Write-Log "Exporting data to $CsvPath"
    $ExportedUsers | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

    Write-Log "--- Export Summary ---"
    Write-Log "Total users found: $TotalUsers"
    Write-Log "Licensed users exported: $LicensedUsersCount"
    Write-Log "Licenses Breakdown:"
    foreach ($Key in $LicenseBreakdown.Keys) {
        Write-Log "  SKU ID $Key : $($LicenseBreakdown[$Key]) users"
    }

} catch {
    Write-Log "An error occurred: $($_.Exception.Message)" -Level ERROR
    throw
} finally {
    Write-Log "Disconnecting from Microsoft Graph..."
    Disconnect-MgGraph
    Write-Log "User Export Process Completed."
}
