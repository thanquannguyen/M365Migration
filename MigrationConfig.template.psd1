<#
==============================================================================
 Microsoft 365 Migration Toolkit - Configuration Template
==============================================================================
 Copy this file to 'MigrationConfig.psd1' and fill in your tenant values.
 'MigrationConfig.psd1' is git-ignored to prevent credential leaks.
#>

@{
    # [Required] Source Tenant *.onmicrosoft.com domain
    SourceTenantDomain  = 'source-tenant.onmicrosoft.com'

    # [Required] Destination Tenant *.onmicrosoft.com domain
    DestTenantDomain    = 'dest-tenant.onmicrosoft.com'

    # [Required] Custom domain for users in the destination tenant (e.g. yourcompany.com)
    NewDomain           = 'destination-domain.com'

    # [Required] Default temporary password for newly created destination users
    DefaultPassword     = 'TemporaryPassword123!Secure'

    # [Required] Target license SKU name (e.g. O365_BUSINESS_ESSENTIALS, ENTERPRISEPACK, STANDARDPACK)
    LicenseSku          = 'O365_BUSINESS_ESSENTIALS'

    # [Required] Azure App Registration Application (client) ID for Source Tenant
    # Permissions required: User.ReadWrite.All, Files.ReadWrite.All, Sites.FullControl.All, Directory.ReadWrite.All
    SourceClientId      = '00000000-0000-0000-0000-000000000000'

    # [Required] Azure App Registration Client Secret value
    SourceClientSecret  = 'YOUR_APPLICATION_SECRET_VALUE_HERE'

    # [Required] Azure App Registration Application (client) ID for Destination Tenant
    # Same permissions as the source app registration, on the destination tenant.
    DestClientId        = '00000000-0000-0000-0000-000000000000'

    # [Required] Azure App Registration Client Secret value (destination tenant)
    DestClientSecret    = 'YOUR_APPLICATION_SECRET_VALUE_HERE'

    # [Optional] Default usage location for license assignment (e.g. VN, US)
    DefaultUsageLocation = 'VN'

    # [Optional] Concurrency: Number of user accounts processed simultaneously (default: 4)
    MaxConcurrentUsers  = 4

    # [Optional] Staging folder for temporary migration files
    StagingPath         = '.\MigrationStaging'

    # [Optional] Output folder for mapping files and reports
    OutputPath          = '.\Output'

    # [Optional] Log directory for detailed execution logs
    LogPath             = '.\Logs'

    # [Optional] Maximum retries for transient Graph API rate limit (429) errors
    MaxRetries          = 3

    # [Optional] Base delay in seconds between retries
    RetryDelaySeconds   = 5

    # [Optional] Chunk upload size in bytes for large OneDrive files (default: 10 MB)
    ChunkSizeBytes      = 10485760

    # [Optional] Force user to change password on first login ($true / $false)
    ForcePasswordChange = $true
}
