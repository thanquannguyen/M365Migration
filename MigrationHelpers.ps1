<#
.SYNOPSIS
    Shared safety helpers for the M365Migration toolkit.

.DESCRIPTION
    Dot-source from any migration script:

        . "$PSScriptRoot\MigrationHelpers.ps1"

    Guarantees enforced by these helpers:
      - No tenant fallback: tenant domains always come from configuration.
        A missing or placeholder tenant value throws instead of silently
        using a hard-coded tenant.
      - No secret is ever persisted by these helpers: access tokens live
        in memory only.
      - Destructive scripts must pass Assert-ValidationGate first: a
        recent passing run of 5-Validate.ps1 is required before any
        purge/delete runs.

    Validation-result file schema (written by 5-Validate.ps1, read by
    Assert-ValidationGate), JSON:
        @{
            Pass          = [bool]
            Timestamp     = [string] ISO 8601 UTC, e.g. "2026-11-08T05:00:00Z"
            TotalUsers    = [int]
            PassedUsers   = [int]
            FailedUsers   = [int]
            FilesExpected = [int]
            FilesActual   = [int]
            FilesMissing  = [int]
        }
#>

function Get-RequiredConfigValue {
    [CmdletBinding()]
    param(
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )
    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw "Configuration value '$Name' is missing or empty. Set it in MigrationConfig.psd1; there is no default tenant or credential fallback."
    }
    if ($Value -eq '00000000-0000-0000-0000-000000000000' -or $Value -like 'YOUR_*') {
        throw "Configuration value '$Name' still holds its template placeholder ('$Value'). Replace it with a real value in MigrationConfig.psd1."
    }
    return $Value
}

function Get-AppOnlyGraphAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantDomain,

        [Parameter(Mandatory = $true)]
        [string]$ClientId,

        [Parameter(Mandatory = $true)]
        [string]$ClientSecret,

        [ValidateRange(1, 10)]
        [int]$Attempts = 4
    )
    $TenantDomain = Get-RequiredConfigValue -Value $TenantDomain -Name 'TenantDomain'
    $ClientId = Get-RequiredConfigValue -Value $ClientId -Name 'ClientId'
    $ClientSecret = Get-RequiredConfigValue -Value $ClientSecret -Name 'ClientSecret'

    $uri = "https://login.microsoftonline.com/$([System.Uri]::EscapeDataString($TenantDomain))/oauth2/v2.0/token"
    $body = @{
        client_id     = $ClientId
        client_secret = $ClientSecret
        scope         = 'https://graph.microsoft.com/.default'
        grant_type    = 'client_credentials'
    }

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        try {
            $response = Invoke-RestMethod -Uri $uri -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded' -TimeoutSec 60 -ErrorAction Stop
            if ([string]::IsNullOrWhiteSpace($response.access_token)) {
                throw 'Token endpoint returned no access_token.'
            }
            return $response.access_token
        }
        catch {
            if ($attempt -ge $Attempts) {
                throw
            }
            $delay = [Math]::Min(30, [Math]::Pow(2, $attempt))
            Write-Warning "Token request failed (attempt $attempt/$Attempts). Retrying in $delay seconds. Error: $($_.Exception.Message)"
            Start-Sleep -Seconds $delay
        }
    }
}

function Assert-Tenant {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExpectedTenantDomain
    )
    $ExpectedTenantDomain = Get-RequiredConfigValue -Value $ExpectedTenantDomain -Name 'ExpectedTenantDomain'

    $context = Get-MgContext
    if ($null -eq $context -or [string]::IsNullOrWhiteSpace($context.TenantId)) {
        throw 'Not connected to Microsoft Graph (Get-MgContext returned nothing). Connect first, then re-run.'
    }

    $org = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id,verifiedDomains'
    $domains = @()
    if ($null -ne $org.value) {
        foreach ($entry in $org.value) {
            foreach ($d in $entry.verifiedDomains) {
                $domains += $d.name
            }
        }
    }

    $matched = $false
    foreach ($d in $domains) {
        if ($d -ieq $ExpectedTenantDomain) {
            $matched = $true
        }
    }
    if (-not $matched) {
        throw "Connected Graph tenant (TenantId '$($context.TenantId)', verified domains: '$($domains -join ', ')') does not match expected tenant domain '$ExpectedTenantDomain'. Aborting: refusing to run against the wrong tenant."
    }
}

function Get-SpoAdminUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantDomain
    )
    $TenantDomain = Get-RequiredConfigValue -Value $TenantDomain -Name 'TenantDomain'
    $tenantName = $TenantDomain -replace '\.onmicrosoft\.com$', ''
    if ([string]::IsNullOrWhiteSpace($tenantName)) {
        throw "Cannot derive a SharePoint admin URL from tenant domain '$TenantDomain'. Expected a '*.onmicrosoft.com' domain."
    }
    return "https://$tenantName-admin.sharepoint.com"
}

function Assert-SpoTenant {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExpectedTenantDomain,

        [Parameter(Mandatory = $true)]
        [string]$AdminUrl
    )
    $ExpectedTenantDomain = Get-RequiredConfigValue -Value $ExpectedTenantDomain -Name 'ExpectedTenantDomain'
    $expectedHost = ([System.Uri](Get-SpoAdminUrl -TenantDomain $ExpectedTenantDomain)).Host.ToLowerInvariant()
    $actualHost = ([System.Uri]$AdminUrl).Host.ToLowerInvariant()
    if ($actualHost -ne $expectedHost) {
        throw "SharePoint admin URL host '$actualHost' does not match expected tenant '$ExpectedTenantDomain' (expected host '$expectedHost'). Aborting: refusing to run against the wrong tenant."
    }
}

function Assert-ValidationGate {
    [CmdletBinding()]
    param(
        [string]$ValidationResultPath = (Join-Path (Join-Path $PSScriptRoot 'Output') 'ValidationResult.json'),

        [ValidateRange(1, 720)]
        [int]$MaxAgeHours = 24
    )
    if (-not (Test-Path -LiteralPath $ValidationResultPath)) {
        throw "Validation gate: no validation result found at '$ValidationResultPath'. Run 5-Validate.ps1 and confirm it passes before any purge/delete."
    }
    $result = Get-Content -LiteralPath $ValidationResultPath -Raw | ConvertFrom-Json
    if ($null -eq $result -or $result.Pass -ne $true) {
        throw "Validation gate: the latest validation result at '$ValidationResultPath' did not pass. Purge/delete is blocked until 5-Validate.ps1 passes."
    }
    $timestamp = [datetime]::MinValue
    if (-not [datetime]::TryParse($result.Timestamp, [ref]$timestamp)) {
        throw "Validation gate: validation result timestamp '$($result.Timestamp)' is not a valid date. Re-run 5-Validate.ps1."
    }
    $age = (Get-Date).ToUniversalTime() - $timestamp.ToUniversalTime()
    if ($age.TotalHours -gt $MaxAgeHours) {
        throw "Validation gate: the latest passing validation is $([math]::Round($age.TotalHours, 1)) hours old (limit: $MaxAgeHours hours). Re-run 5-Validate.ps1 before any purge/delete."
    }
    Write-Host "Validation gate OK: passing validation from $($result.Timestamp) (age $([math]::Round($age.TotalHours, 1))h)." -ForegroundColor Green
}

function Invoke-WithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [ValidateRange(1, 10)]
        [int]$MaxAttempts = 3,

        [ValidateRange(0, 300)]
        [int]$BaseDelaySeconds = 5
    )
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return & $ScriptBlock
        }
        catch {
            if ($attempt -ge $MaxAttempts) {
                throw
            }
            $delay = [Math]::Min(300, $BaseDelaySeconds * [Math]::Pow(2, $attempt - 1))
            Write-Warning "Attempt $attempt/$MaxAttempts failed: $($_.Exception.Message). Retrying in $delay seconds."
            Start-Sleep -Seconds $delay
        }
    }
}
