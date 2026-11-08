<#
.SYNOPSIS
    Validates user migration and generates a report.

    Requires (minimum Graph scopes): User.Read.All, Files.Read.All,
    Sites.Read.All, Organization.Read.All, Directory.Read.All.

    Writes Output/ValidationResult.json (the gate read by
    Assert-ValidationGate before any purge/delete script runs).
#>
[CmdletBinding()]
param (
    [string]$ConfigPath = '.\MigrationConfig.psd1',
    [string]$MappingPath = '.\Output\UserMapping.csv'
)

. "$PSScriptRoot\MigrationHelpers.ps1"

$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"
$env:AZURE_IDENTITY_DISABLE_WAM = "true"

$LogDir = ".\Logs"
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
$LogFile = Join-Path $LogDir "Validate_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param([string]$Message, [string]$Color = "Cyan")
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] $Message"
    Add-Content -Path $LogFile -Value $LogMessage
    Write-Host $LogMessage -ForegroundColor $Color
}

function Write-ValidationResult {
    param(
        [bool]$Pass,
        [long]$TotalUsers,
        [long]$PassedUsers,
        [long]$FailedUsers,
        [long]$FilesExpected,
        [long]$FilesActual,
        [long]$FilesMissing
    )
    $GateDir = ".\Output"
    if (-not (Test-Path $GateDir)) { New-Item -ItemType Directory -Path $GateDir -Force | Out-Null }
    $result = [ordered]@{
        Pass          = $Pass
        Timestamp     = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        TotalUsers    = $TotalUsers
        PassedUsers   = $PassedUsers
        FailedUsers   = $FailedUsers
        FilesExpected = $FilesExpected
        FilesActual   = $FilesActual
        FilesMissing  = $FilesMissing
    }
    $result | ConvertTo-Json | Set-Content -Path (Join-Path $GateDir 'ValidationResult.json') -Encoding UTF8
}

function Get-DestDriveStats {
    # S3: RECURSIVE destination count + total bytes via delta paging.
    param([string]$DriveId)
    $count = 0
    $bytes = [long]0
    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root/delta?`$select=id,name,size,folder,file,deleted"
    try {
        while ($uri) {
            $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            foreach ($item in $resp.value) {
                if ($item.deleted) { continue }
                if ($null -ne $item.file) {
                    $count++
                    $bytes += [long]$item.size
                }
            }
            $uri = $resp.'@odata.nextLink'
        }
        return [PSCustomObject]@{ Count = $count; Bytes = $bytes; Error = $null }
    } catch {
        return [PSCustomObject]@{ Count = $count; Bytes = $bytes; Error = $_.Exception.Message }
    }
}

function Get-ManifestExpected {
    # S3: manifest.json is the source of truth for expected files. Script 4
    # pre-populates it with EVERY source file (Id/Path/Size/Downloaded/
    # Uploaded/Skipped) and 4b appends uploaded FileIds to uploaded.log;
    # uploaded.log alone only lists already-uploaded ids, so the manifest
    # (not uploaded.log, not a staging recount) defines the expected total.
    # Permanently-skipped entries (deleted/blocked at source) are excluded:
    # they can never be migrated. Falls back to a recursive staging count
    # only when no manifest exists; returns Found=$false when neither exists.
    param([string]$StagingDir)
    $manifestPath = Join-Path $StagingDir 'manifest.json'
    if (Test-Path $manifestPath) {
        try {
            $json = Get-Content $manifestPath -Raw -ErrorAction Stop | ConvertFrom-Json
            $entries = @($json.PSObject.Properties | ForEach-Object { $_.Value } | Where-Object { $_.Skipped -ne $true })
            $totalBytes = ($entries | Measure-Object -Property Size -Sum).Sum
            if ($null -eq $totalBytes) { $totalBytes = 0 }
            return [PSCustomObject]@{ Source = 'manifest.json'; Count = $entries.Count; Bytes = [long]$totalBytes; Found = $true }
        } catch {
            Write-Log "  WARNING: manifest.json at $manifestPath is unreadable ($($_.Exception.Message)); falling back to staging count." "Yellow"
        }
    }
    if (Test-Path $StagingDir) {
        $files = @(Get-ChildItem -Path $StagingDir -File -Recurse -Force)
        $totalBytes = ($files | Measure-Object -Property Length -Sum).Sum
        if ($null -eq $totalBytes) { $totalBytes = 0 }
        return [PSCustomObject]@{ Source = 'staging'; Count = $files.Count; Bytes = [long]$totalBytes; Found = $true }
    }
    return [PSCustomObject]@{ Source = 'none'; Count = 0; Bytes = 0; Found = $false }
}

Write-Log "Starting Validation process" "Green"

if (-not (Test-Path $MappingPath)) {
    Write-Log "Mapping file not found at $MappingPath" "Red"
    Write-ValidationResult -Pass $false -TotalUsers 0 -PassedUsers 0 -FailedUsers 0 -FilesExpected 0 -FilesActual 0 -FilesMissing 0
    return
}

$Mapping = Import-Csv $MappingPath
$ReportData = @()

# Load config for tenant info
$config = @{}
if (Test-Path $ConfigPath) {
    $config = Import-PowerShellDataFile -Path $ConfigPath
}

Write-Log "Connecting to Microsoft Graph (Device Code)..." "Yellow"
Write-Log "A device code will appear below. Go to https://microsoft.com/devicelogin and enter it." "Yellow"
Write-Log "Sign in with your DESTINATION tenant admin account." "Yellow"
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    # No fallback: the destination tenant must come from config.
    $destTenantId = Get-RequiredConfigValue -Value $config.DestTenantDomain -Name 'DestTenantDomain'
    $valScopes = @('User.Read.All', 'Files.Read.All', 'Sites.Read.All', 'Organization.Read.All', 'Directory.Read.All')
    Connect-MgGraph -Scopes $valScopes -TenantId $destTenantId -UseDeviceCode -NoWelcome *>&1 | Write-Host
    # B1: refuse to validate against the wrong tenant.
    Assert-Tenant -ExpectedTenantDomain $destTenantId
    Write-Log "Connected to destination tenant successfully." "Green"
} catch {
    Write-Log "Failed to connect to Graph: $_" "Red"
    Write-ValidationResult -Pass $false -TotalUsers 0 -PassedUsers 0 -FailedUsers 0 -FilesExpected 0 -FilesActual 0 -FilesMissing 0
    return
}

foreach ($User in $Mapping) {
    $TargetUPN = $User.DestUPN
    $SourceUPN = $User.SourceUPN
    Write-Log "Validating user: $TargetUPN" "Cyan"
    
    $UserStatus = "Missing"
    $LicenseStatus = "Unlicensed"
    $FilesExpected = 0
    $FilesActual = 0
    $BytesExpected = [long]0
    $BytesActual = [long]0
    $ExpectedSource = 'none'
    $ExpectedUnknown = $false
    
    # Check User
    try {
        $MgUser = Get-MgUser -UserId $TargetUPN -ErrorAction Stop
        $UserStatus = "Exists"
        
        # Check License
        $Licenses = Get-MgUserLicenseDetail -UserId $TargetUPN -ErrorAction Stop
        if ($Licenses -and $Licenses.Count -gt 0) {
            $LicenseStatus = "Licensed"
        }
        
        # Check OneDrive files: RECURSIVE count + total bytes via delta (S3).
        # The old code counted only the drive root, under-counting every
        # user with subfolders.
        try {
            $Drive = Get-MgUserDrive -UserId $TargetUPN -ErrorAction Stop
            if ($Drive) {
                $destStats = Get-DestDriveStats -DriveId $Drive.Id
                if ($destStats.Error) {
                    Write-Log "  Could not enumerate OneDrive for ${TargetUPN}: $($destStats.Error)" "Yellow"
                }
                $FilesActual = $destStats.Count
                $BytesActual = $destStats.Bytes
            }
        } catch {
            Write-Log "  Could not access OneDrive for ${TargetUPN}: $_" "Yellow"
        }
    } catch {
        Write-Log "  User $TargetUPN not found." "Yellow"
    }
    
    # Expected files: manifest.json is the source of truth (S3). The upload
    # worker deletes staging files after upload, so counting staging alone
    # would report FilesExpected=0 as a false success. If NEITHER the
    # manifest nor the staging folder exists, the expected count is UNKNOWN
    # and this user cannot pass validation.
    $StagingDir = ".\MigrationStaging\$SourceUPN"
    $expectedInfo = Get-ManifestExpected -StagingDir $StagingDir
    $FilesExpected = $expectedInfo.Count
    $BytesExpected = $expectedInfo.Bytes
    $ExpectedSource = $expectedInfo.Source
    if (-not $expectedInfo.Found) {
        $ExpectedUnknown = $true
        Write-Log "  WARNING: no manifest.json and no staging folder for $SourceUPN - expected file count is UNKNOWN, user cannot pass validation." "Red"
    } elseif ($expectedInfo.Source -eq 'staging') {
        Write-Log "  WARNING: no manifest.json for $SourceUPN - expected count from staging folder only (may under-count if the worker already cleaned files)." "Yellow"
    }
    
    $FilesMissing = $FilesExpected - $FilesActual
    if ($FilesMissing -lt 0) { $FilesMissing = 0 }
    # Size check (S3): signed diff for the report; missing bytes (expected
    # minus actual, floored at 0) for the pass/fail decision. Extra bytes on
    # dest (e.g. announcement files) only warn, they do not fail the user.
    $SizeMismatchBytes = $BytesActual - $BytesExpected
    $SizeMissingBytes = $BytesExpected - $BytesActual
    if ($SizeMissingBytes -lt 0) { $SizeMissingBytes = 0 }

    $ReportData += [PSCustomObject]@{
        TargetUPN        = $TargetUPN
        UserStatus       = $UserStatus
        LicenseStatus    = $LicenseStatus
        FilesExpected    = $FilesExpected
        FilesActual      = $FilesActual
        FilesMissing     = $FilesMissing
        BytesExpected    = $BytesExpected
        BytesActual      = $BytesActual
        SizeMismatchBytes = $SizeMismatchBytes
        SizeMissingBytes = $SizeMissingBytes
        ExpectedSource   = $ExpectedSource
    }
}

Write-Log "Disconnecting from Graph..." "Yellow"
Disconnect-MgGraph

# Generate HTML Report
$OutputDir = ".\Output"
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
$ReportPath = Join-Path $OutputDir "MigrationReport.html"
Write-Log "Generating report at $ReportPath" "Cyan"

# S3: a user only counts as successful when the account exists, is licensed
# AND every expected file is present (file completeness) with no missing
# bytes. An unknown expected count (manifest + staging both gone) never passes.
$SuccessCount = ($ReportData | Where-Object { $_.UserStatus -eq 'Exists' -and $_.LicenseStatus -eq 'Licensed' -and $_.FilesMissing -eq 0 -and $_.SizeMissingBytes -eq 0 -and $_.ExpectedSource -ne 'none' }).Count
$TotalUsers = $ReportData.Count
$SuccessRate = if ($TotalUsers -gt 0) { [math]::Round(($SuccessCount / $TotalUsers) * 100, 2) } else { 0 }

$HtmlHead = @"
<style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #1e1e1e; color: #d4d4d4; margin: 20px; }
    h1, h2 { color: #569cd6; }
    table { border-collapse: collapse; width: 100%; margin-top: 20px; }
    th, td { border: 1px solid #333; padding: 10px; text-align: left; }
    th { background-color: #252526; color: #4ec9b0; }
    tr:nth-child(even) { background-color: #2d2d30; }
    .good { color: #89d185; font-weight: bold; }
    .issue { color: #f48771; font-weight: bold; }
    .summary-card { background-color: #252526; padding: 15px; border-radius: 5px; margin-bottom: 20px; border-left: 4px solid #569cd6; }
</style>
"@

$HtmlBody = "<h1>Microsoft 365 Migration Report</h1>"
$HtmlBody += "<div class='summary-card'>"
$HtmlBody += "<p><strong>Timestamp:</strong> $(Get-Date)</p>"
$HtmlBody += "<p><strong>Total Users:</strong> $TotalUsers</p>"
$HtmlBody += "<p><strong>Fully Successful Users (Exists & Licensed & Files Complete):</strong> $SuccessCount</p>"
$HtmlBody += "<p><strong>Success Rate:</strong> $SuccessRate%</p>"
$HtmlBody += "</div>"

$HtmlBody += "<h2>Per-User Details</h2>"
$HtmlBody += "<table><tr><th>Target UPN</th><th>User Status</th><th>License Status</th><th>Expected Files</th><th>Actual Files</th><th>Missing Files</th><th>Expected Bytes</th><th>Actual Bytes</th><th>Size Diff (Bytes)</th><th>Expected From</th></tr>"

foreach ($Row in $ReportData) {
    $UClass = if ($Row.UserStatus -eq 'Exists') { "good" } else { "issue" }
    $LClass = if ($Row.LicenseStatus -eq 'Licensed') { "good" } else { "issue" }
    $FClass = if ($Row.FilesMissing -eq 0) { "good" } else { "issue" }
    
    $HtmlBody += "<tr>"
    $HtmlBody += "<td>$($Row.TargetUPN)</td>"
    $HtmlBody += "<td class='$UClass'>$($Row.UserStatus)</td>"
    $HtmlBody += "<td class='$LClass'>$($Row.LicenseStatus)</td>"
    $SClass = if ($Row.SizeMissingBytes -eq 0) { "good" } else { "issue" }
    $HtmlBody += "<td>$($Row.FilesExpected)</td>"
    $HtmlBody += "<td>$($Row.FilesActual)</td>"
    $HtmlBody += "<td class='$FClass'>$($Row.FilesMissing)</td>"
    $HtmlBody += "<td>$($Row.BytesExpected)</td>"
    $HtmlBody += "<td>$($Row.BytesActual)</td>"
    $HtmlBody += "<td class='$SClass'>$($Row.SizeMismatchBytes)</td>"
    $HtmlBody += "<td>$($Row.ExpectedSource)</td>"
    $HtmlBody += "</tr>"
}

$HtmlBody += "</table>"

$Html = "<html><head>$HtmlHead</head><body>$HtmlBody</body></html>"
$Html | Out-File -FilePath $ReportPath -Encoding UTF8

# Write the validation gate file consumed by Assert-ValidationGate (B2).
$PassedUsers = $SuccessCount
$FailedUsers = $TotalUsers - $SuccessCount
$GatePass = ($TotalUsers -gt 0 -and $FailedUsers -eq 0)
$TotalFilesExpected = ($ReportData | Measure-Object -Property FilesExpected -Sum).Sum
$TotalFilesActual = ($ReportData | Measure-Object -Property FilesActual -Sum).Sum
$TotalFilesMissing = ($ReportData | Measure-Object -Property FilesMissing -Sum).Sum
if ($null -eq $TotalFilesExpected) { $TotalFilesExpected = 0 }
if ($null -eq $TotalFilesActual) { $TotalFilesActual = 0 }
if ($null -eq $TotalFilesMissing) { $TotalFilesMissing = 0 }
Write-ValidationResult -Pass $GatePass -TotalUsers $TotalUsers -PassedUsers $PassedUsers -FailedUsers $FailedUsers -FilesExpected $TotalFilesExpected -FilesActual $TotalFilesActual -FilesMissing $TotalFilesMissing
if ($GatePass) {
    Write-Log "Validation PASSED: all $TotalUsers user(s) fully successful. Gate file written to Output/ValidationResult.json." "Green"
} else {
    Write-Log "Validation FAILED: $FailedUsers of $TotalUsers user(s) did not fully pass. Gate file written to Output/ValidationResult.json (Pass=false); purge scripts are blocked until this passes." "Red"
}

Write-Log "Validation completed." "Green"

# Console Summary
Write-Host "`n--- Migration Summary ---" -ForegroundColor Cyan
Write-Host "Total Users: $TotalUsers"
Write-Host "Passed Users: $PassedUsers | Failed Users: $FailedUsers"
Write-Host "Success Rate: $SuccessRate%"
Write-Host "Files Expected: $TotalFilesExpected | Files Actual: $TotalFilesActual | Files Missing: $TotalFilesMissing"
$ReportData | Format-Table -AutoSize
