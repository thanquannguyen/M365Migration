<#
.SYNOPSIS
    Creates users in the destination Microsoft 365 tenant using the Microsoft Graph PowerShell SDK.

.DESCRIPTION
    This script reads exported users from a CSV file and creates them in a destination tenant.
    It maps the old domain to the new domain, sets a default password, and handles existing users.
    It outputs a user mapping file and a user list WITHOUT passwords: no password is ever
    written to disk (B4). New users share the DefaultPassword from the configuration file
    and must change it at first sign-in (ForceChangePasswordNextSignIn comes from config).

    Requires (minimum Graph scopes): User.ReadWrite.All, Directory.ReadWrite.All, Organization.Read.All (delegated, interactive sign-in; Organization.Read.All is needed by the tenant check).

.PARAMETER ConfigPath
    Path to the migration configuration file. Default is '.\MigrationConfig.psd1'.

.PARAMETER CsvPath
    Path to the CSV file containing exported users. Default is '.\Output\ExportedUsers.csv'.

.PARAMETER NewDomain
    Optional parameter to override the TargetTenantDomain from the configuration file.

.EXAMPLE
    .\2-CreateUsers.ps1 -ConfigPath '.\MigrationConfig.psd1' -CsvPath '.\Output\ExportedUsers.csv' -NewDomain 'newdomain.com'
#>

[CmdletBinding()]
param (
    [string]$ConfigPath = '.\MigrationConfig.psd1',
    [string]$CsvPath = '.\Output\ExportedUsers.csv',
    [string]$NewDomain
)

. "$PSScriptRoot\MigrationHelpers.ps1"

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'
$env:AZURE_IDENTITY_DISABLE_WAM = 'true'

# Setup Logging
$logDir = '.\Logs'
if (!(Test-Path -Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir | Out-Null
}
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logFile = Join-Path -Path $logDir -ChildPath "CreateUsers_$timestamp.log"

function Write-Log {
    param (
        [string]$Message,
        [string]$Level = 'INFO'
    )
    $logMessage = "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) [$Level] $Message"
    Add-Content -Path $logFile -Value $logMessage
    
    switch ($Level) {
        'INFO' { Write-Host $Message -ForegroundColor Cyan }
        'SUCCESS' { Write-Host $Message -ForegroundColor Green }
        'WARNING' { Write-Host $Message -ForegroundColor Yellow }
        'ERROR' { Write-Host $Message -ForegroundColor Red }
    }
}

Write-Log -Message "Starting user creation process..." -Level 'INFO'

# Load Configuration
if (!(Test-Path -Path $ConfigPath)) {
    Write-Log -Message "Configuration file not found at $ConfigPath" -Level 'ERROR'
    return
}

try {
    $config = Import-PowerShellDataFile -Path $ConfigPath
    Write-Log -Message "Successfully loaded configuration from $ConfigPath" -Level 'INFO'
} catch {
    Write-Log -Message "Failed to load configuration: $_" -Level 'ERROR'
    return
}

# Determine New Domain
$targetDomain = $NewDomain
if ([string]::IsNullOrWhiteSpace($targetDomain)) {
    $targetDomain = $config.NewDomain
}

if ([string]::IsNullOrWhiteSpace($targetDomain)) {
    Write-Log -Message "Target domain not specified in parameters or config." -Level 'ERROR'
    return
}
Write-Log -Message "Using target domain: $targetDomain" -Level 'INFO'

# Read CSV
if (!(Test-Path -Path $CsvPath)) {
    Write-Log -Message "CSV file not found at $CsvPath" -Level 'ERROR'
    return
}

try {
    $users = Import-Csv -Path $CsvPath -Encoding UTF8
    Write-Log -Message "Successfully loaded $($users.Count) users from $CsvPath" -Level 'INFO'
} catch {
    Write-Log -Message "Failed to load CSV: $_" -Level 'ERROR'
    return
}

# Ensure Output directory exists
$outputDir = '.\Output'
if (!(Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

try {
    Write-Log -Message "Connecting to destination tenant..." -Level 'INFO'
    Write-Log -Message "A browser window will open. Sign in with your DESTINATION tenant admin account." -Level 'WARNING'
    $connectParams = @{ Scopes = @('User.ReadWrite.All','Directory.ReadWrite.All','Organization.Read.All') }
    $destTenantDomain = Get-RequiredConfigValue -Value $config.DestTenantDomain -Name 'DestTenantDomain'
    $connectParams.TenantId = $destTenantDomain
    Connect-MgGraph @connectParams -ErrorAction Stop
    Assert-Tenant -ExpectedTenantDomain $destTenantDomain
    Write-Log -Message "Connected successfully. Tenant verified against DestTenantDomain." -Level 'SUCCESS'
} catch {
    Write-Log -Message "Failed to connect to Microsoft Graph: $_" -Level 'ERROR'
    return
}

$userMappings = @()
$userCredentials = @()
$stats = @{
    Created = 0
    Skipped = 0
    Failed = 0
}

# S6/B4: DefaultPassword is mandatory from config (no weak fallback);
# ForceChangePasswordNextSignIn comes from config, defaulting to $true.
$defaultPassword = Get-RequiredConfigValue -Value $config.DefaultPassword -Name 'DefaultPassword'
$forcePasswordChange = if ($null -ne $config.ForcePasswordChange) { [bool]$config.ForcePasswordChange } else { $true }

foreach ($user in $users) {
    try {
        $sourceUpn = $user.UserPrincipalName
        
        # Extract username part
        $username = $sourceUpn.Split('@')[0]
        $destUpn = "$username@$targetDomain"
        
        Write-Log -Message "Processing user: $sourceUpn -> $destUpn" -Level 'INFO'
        
        # Check if user already exists (UPN escaped for the OData -Filter string)
        $escapedDestUpn = $destUpn -replace "'", "''"
        $existingUser = Get-MgUser -Filter "userPrincipalName eq '$escapedDestUpn'" -ErrorAction SilentlyContinue
        
        if ($existingUser) {
            Write-Log -Message "User $destUpn already exists in destination. Skipping." -Level 'WARNING'
            $stats.Skipped++
            
            $userMappings += [PSCustomObject]@{
                SourceUPN = $sourceUpn
                DestUPN = $destUpn
                SourceId = $user.Id
                DestId = $existingUser.Id
                Status = 'Skipped'
            }
            continue
        }
        
        # Prepare parameters for new user
        $usageLocation = $user.UsageLocation
        if ([string]::IsNullOrWhiteSpace($usageLocation)) {
            $usageLocation = 'VN'
        }
        
        # Format BusinessPhones as array: split on ';', trim, drop empties (S6)
        $businessPhones = @()
        if (![string]::IsNullOrWhiteSpace($user.BusinessPhones)) {
            $businessPhones = @($user.BusinessPhones -split ';' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        }

        $newUserParams = @{
            DisplayName = $user.DisplayName
            MailNickname = $username
            UserPrincipalName = $destUpn
            UsageLocation = $usageLocation
            AccountEnabled = $true
            PasswordProfile = @{
                Password = $defaultPassword
                ForceChangePasswordNextSignIn = $forcePasswordChange
            }
        }
        
        # Add optional properties if they exist
        if (![string]::IsNullOrWhiteSpace($user.GivenName)) { $newUserParams.Add('GivenName', $user.GivenName) }
        if (![string]::IsNullOrWhiteSpace($user.Surname)) { $newUserParams.Add('Surname', $user.Surname) }
        if (![string]::IsNullOrWhiteSpace($user.JobTitle)) { $newUserParams.Add('JobTitle', $user.JobTitle) }
        if (![string]::IsNullOrWhiteSpace($user.Department)) { $newUserParams.Add('Department', $user.Department) }
        if (![string]::IsNullOrWhiteSpace($user.OfficeLocation)) { $newUserParams.Add('OfficeLocation', $user.OfficeLocation) }
        if (![string]::IsNullOrWhiteSpace($user.CompanyName)) { $newUserParams.Add('CompanyName', $user.CompanyName) }
        if (![string]::IsNullOrWhiteSpace($user.City)) { $newUserParams.Add('City', $user.City) }
        if (![string]::IsNullOrWhiteSpace($user.Country)) { $newUserParams.Add('Country', $user.Country) }
        if (![string]::IsNullOrWhiteSpace($user.PostalCode)) { $newUserParams.Add('PostalCode', $user.PostalCode) }
        if (![string]::IsNullOrWhiteSpace($user.State)) { $newUserParams.Add('State', $user.State) }
        if (![string]::IsNullOrWhiteSpace($user.StreetAddress)) { $newUserParams.Add('StreetAddress', $user.StreetAddress) }
        if (![string]::IsNullOrWhiteSpace($user.MobilePhone)) { $newUserParams.Add('MobilePhone', $user.MobilePhone) }
        if ($businessPhones.Count -gt 0) { $newUserParams.Add('BusinessPhones', $businessPhones) }

        # Create user
        $newUser = New-MgUser @newUserParams -ErrorAction Stop
        
        Write-Log -Message "Successfully created user $destUpn" -Level 'SUCCESS'
        $stats.Created++
        
        $userMappings += [PSCustomObject]@{
            SourceUPN = $sourceUpn
            DestUPN = $destUpn
            SourceId = $user.Id
            DestId = $newUser.Id
            Status = 'Created'
        }
        
        # B4: passwords are NEVER written to disk. The user list below carries
        # no password column; the admin hands out the configured DefaultPassword
        # and each user must change it at first sign-in.
        $userCredentials += [PSCustomObject]@{
            DisplayName = $newUser.DisplayName
            NewUPN = $destUpn
        }
        
    } catch {
        Write-Log -Message "Failed to create user $sourceUpn : $_" -Level 'ERROR'
        $stats.Failed++
        
        $userMappings += [PSCustomObject]@{
            SourceUPN = $sourceUpn
            DestUPN = $destUpn
            SourceId = $user.Id
            DestId = ''
            Status = 'Failed'
        }
    }
}

# Export mapping and credentials files
try {
    $mappingPath = Join-Path -Path $outputDir -ChildPath 'UserMapping.csv'
    $userMappings | Export-Csv -Path $mappingPath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "Exported user mapping to $mappingPath" -Level 'INFO'
    
    $credPath = Join-Path -Path $outputDir -ChildPath 'UserCredentials.csv'
    $userCredentials | Export-Csv -Path $credPath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "Exported user list (no passwords) to $credPath" -Level 'INFO'
    Write-Host "Passwords are NOT stored in $credPath or anywhere on disk. New users share the DefaultPassword from the configuration file and must change it at first sign-in." -ForegroundColor Yellow
} catch {
    Write-Log -Message "Failed to export output CSV files: $_" -Level 'ERROR'
}

Write-Log -Message "Summary: Created: $($stats.Created) | Skipped: $($stats.Skipped) | Failed: $($stats.Failed)" -Level 'INFO'
Write-Log -Message "User creation process completed." -Level 'INFO'

# Disconnect
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Write-Log -Message "Disconnected from Microsoft Graph." -Level 'INFO'
} catch {
    Write-Log -Message "Failed to disconnect from Microsoft Graph cleanly." -Level 'WARNING'
}
