<#
.SYNOPSIS
Renames all CHU_Y.txt announcement files to .CHU_Y.txt on source tenant OneDrive accounts.

Requires (minimum Graph scopes): Files.ReadWrite.All, User.Read.All
(app-only, via client credentials / https://graph.microsoft.com/.default).

.DESCRIPTION
Leading dot (.CHU_Y.txt) ensures the announcement file sorts to the very top in Windows File Explorer
and OneDrive web interface, making it immediately visible to users.

.PARAMETER ConfigPath
Path to MigrationConfig.psd1.

.PARAMETER MappingPath
Path to Output\UserMapping.csv.

.PARAMETER UserFilter
Optional regex or username filter to target specific accounts.
#>

# S13: standard confirmation (SupportsShouldProcess). ConfirmImpact Medium
# so a plain run proceeds without prompting per file, while -WhatIf previews
# and -Confirm prompts before each rename/delete.
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param (
    [string]$ConfigPath = '.\MigrationConfig.psd1',
    [string]$MappingPath = '.\Output\UserMapping.csv',
    [string]$UserFilter
)

. "$PSScriptRoot\MigrationHelpers.ps1"

$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

try {
    [System.Net.ServicePointManager]::DefaultConnectionLimit = 500
    [System.Net.ServicePointManager]::Expect100Continue = $false
    [System.Net.ServicePointManager]::UseNagleAlgorithm = $false
} catch {}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
if ([string]::IsNullOrWhiteSpace($ScriptDir)) {
    $ScriptDir = (Get-Location).Path
}
Set-Location -Path $ScriptDir

# Setup Logging
$LogDir = Join-Path -Path $ScriptDir -ChildPath 'Logs'
if (-not (Test-Path -Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogSuffix = if ($UserFilter) { "_" + ($UserFilter -replace '[^a-zA-Z0-9]', '_') } else { "_master" }
$LogPath = Join-Path -Path $LogDir -ChildPath "RenameAnnouncement_${Timestamp}${LogSuffix}_${PID}.log"

function Write-Log {
    param ([string]$Message, [string]$Type = "INFO")
    $LogMsg = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Type] $Message"
    $color = 'Cyan'
    if ($Type -eq 'ERROR') { $color = 'Red' }
    elseif ($Type -eq 'WARN') { $color = 'Yellow' }
    elseif ($Type -eq 'SUCCESS') { $color = 'Green' }
    
    Write-Host $LogMsg -ForegroundColor $color
    Add-Content -Path $LogPath -Value $LogMsg -Encoding UTF8 -ErrorAction SilentlyContinue
}

# Load Configuration
if (-not (Test-Path $ConfigPath)) {
    $ConfigPath = Join-Path $ScriptDir 'MigrationConfig.psd1'
}
$config = @{}
if (Test-Path $ConfigPath) {
    $config = Import-PowerShellDataFile -Path $ConfigPath
}

# Load Mapping
if (-not (Test-Path $MappingPath)) {
    $MappingPath = Join-Path $ScriptDir 'Output\UserMapping.csv'
}
if (-not (Test-Path $MappingPath)) {
    Write-Log "Mapping file not found at $MappingPath" -Type ERROR
    exit 1
}
$Mapping = @(Import-Csv -Path $MappingPath -Encoding UTF8)

if ($UserFilter) {
    $exactMatches = @($Mapping | Where-Object { 
        $_.SourceUPN -eq $UserFilter -or 
        $_.SourceUPN -like "$UserFilter@*" -or 
        $_.DestUPN -eq $UserFilter -or 
        $_.DestUPN -like "$UserFilter@*" 
    })
    if ($exactMatches.Count -gt 0) {
        $Mapping = $exactMatches
    } else {
        $Mapping = @($Mapping | Where-Object { $_.SourceUPN -match $UserFilter -or $_.DestUPN -match $UserFilter })
    }
    Write-Log "Filtered to user(s): $UserFilter ($($Mapping.Count) user(s))"
}

Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host "   RENAME ANNOUNCEMENT FILES TO .CHU_Y.txt" -ForegroundColor Cyan
Write-Host "   Target: Files starting with '.' appear at the TOP in File Explorer" -ForegroundColor Yellow
Write-Host "   Total Accounts: $($Mapping.Count)" -ForegroundColor Cyan
Write-Host "==========================================================`n" -ForegroundColor Cyan

# OAuth Token Management
# No fallback: tenant/credentials must come from config.
$tenantId = Get-RequiredConfigValue -Value $config.SourceTenantDomain -Name 'SourceTenantDomain'
$clientId = Get-RequiredConfigValue -Value $config.SourceClientId -Name 'SourceClientId'
$clientSecret = Get-RequiredConfigValue -Value $config.SourceClientSecret -Name 'SourceClientSecret'

$script:cachedToken = $null
$script:tokenExpiry = [DateTime]::MinValue

function Get-SourceGraphToken {
    if ($script:cachedToken -and [DateTime]::UtcNow.AddMinutes(2) -lt $script:tokenExpiry) {
        return $script:cachedToken
    }
    
    $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    $body = @{
        client_id     = $clientId
        client_secret = $clientSecret
        scope         = 'https://graph.microsoft.com/.default'
        grant_type    = 'client_credentials'
    }
    
    $resp = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body
    $script:cachedToken = $resp.access_token
    $script:tokenExpiry = [DateTime]::UtcNow.AddSeconds([int]$resp.expires_in)
    return $script:cachedToken
}

function Invoke-GraphApiWithRetry {
    param (
        [string]$Uri,
        [string]$Method = 'GET',
        [string]$Body = $null,
        [int]$MaxRetries = 5
    )
    
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        $token = Get-SourceGraphToken
        $headers = @{
            Authorization  = "Bearer $token"
            "Content-Type" = "application/json"
        }
        try {
            $params = @{
                Uri     = $Uri
                Method  = $Method
                Headers = $headers
            }
            if ($Body) { $params['Body'] = $Body }
            return (Invoke-RestMethod @params)
        } catch {
            $status = 0
            if ($_.Exception.Response) {
                try { $status = [int]$_.Exception.Response.StatusCode } catch {}
            }
            if ($status -eq 404) {
                return $null # Not found is expected for missing files/folders
            }
            if ($status -eq 429 -or $status -ge 500) {
                $retryAfter = 2
                if ($attempt -lt $MaxRetries) {
                    Start-Sleep -Seconds ($retryAfter * $attempt)
                    continue
                }
            }
            throw $_
        }
    }
}

# S13: compare the two announcement files' contents. Returns $true only when
# they are provably identical; any doubt (sizes differ, hashes differ,
# download/compare fails) returns $false so BOTH files are kept.
function Test-AnnouncementContentsEqual {
    param ([string]$DriveId, $OldItem, $NewItem)
    if (-not $OldItem -or -not $NewItem) { return $false }
    if ($OldItem.size -ne $NewItem.size) { return $false }
    $oldHash = $OldItem.file.hashes.sha1Hash
    $newHash = $NewItem.file.hashes.sha1Hash
    if ($oldHash -and $newHash) { return ($oldHash -ceq $newHash) }
    try {
        $oldContent = Invoke-GraphApiWithRetry -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$($OldItem.id)/content"
        $newContent = Invoke-GraphApiWithRetry -Uri "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$($NewItem.id)/content"
        if ($null -eq $oldContent -or $null -eq $newContent) { return $false }
        return ([string]::Equals("$oldContent", "$newContent", [System.StringComparison]::Ordinal))
    } catch {
        return $false
    }
}

$summary = @()

foreach ($User in $Mapping) {
    $sourceUpn = $User.SourceUPN
    Write-Log "--------------------------------------------------"
    Write-Log "Processing: $sourceUpn"
    
    # 1. Get Drive
    $drive = $null
    try {
        $drive = Invoke-GraphApiWithRetry -Uri "https://graph.microsoft.com/v1.0/users/$sourceUpn/drive"
    } catch {
        Write-Log "  Could not access drive for ${sourceUpn}: $_" -Type ERROR
        $summary += [PSCustomObject]@{
            SourceUPN      = $sourceUpn
            RenamedCount   = 0
            AlreadyRenamed = 0
            Status         = "Failed - No Drive"
        }
        continue
    }
    
    $driveId = $drive.id
    
    # 2. Build list of candidate folder paths
    # Known announcement folders: root, Desktop, Documents, Pictures
    $targetFolders = [System.Collections.Generic.List[string]]::new()
    $targetFolders.Add("root")
    $targetFolders.Add("Desktop")
    $targetFolders.Add("Documents")
    $targetFolders.Add("Pictures")
    
    # Also inspect all direct child folders of root to discover any other folders
    try {
        $rootChildren = Invoke-GraphApiWithRetry -Uri "https://graph.microsoft.com/v1.0/drives/$driveId/root/children"
        if ($rootChildren -and $rootChildren.value) {
            foreach ($item in $rootChildren.value) {
                if ($item.folder -and -not $targetFolders.Contains($item.name)) {
                    $targetFolders.Add($item.name)
                }
            }
        }
    } catch {
        Write-Log "  Could not list root children for ${sourceUpn}: $_" -Type WARN
    }
    
    $renamedCount = 0
    $alreadyCount = 0
    
    # 3. Check each target folder for CHU_Y.txt and .CHU_Y.txt
    foreach ($folderName in $targetFolders) {
        # S13: encode the folder name so names with '#', '?', '%' etc. survive.
        $oldFilePath = if ($folderName -eq "root") { "root:/CHU_Y.txt" } else { "root:/$([System.Uri]::EscapeDataString($folderName))/CHU_Y.txt" }
        $newFilePath = if ($folderName -eq "root") { "root:/.CHU_Y.txt" } else { "root:/$([System.Uri]::EscapeDataString($folderName))/.CHU_Y.txt" }
        
        $oldItem = $null
        $newItem = $null
        
        try {
            $oldItem = Invoke-GraphApiWithRetry -Uri "https://graph.microsoft.com/v1.0/drives/$driveId/$oldFilePath"
        } catch {}
        
        try {
            $newItem = Invoke-GraphApiWithRetry -Uri "https://graph.microsoft.com/v1.0/drives/$driveId/$newFilePath"
        } catch {}
        
        # Case 1: .CHU_Y.txt already exists
        if ($newItem) {
            $alreadyCount++
            if ($oldItem) {
                # S13: only delete the old file when its content is IDENTICAL
                # to the new one; if they differ, keep both and warn.
                if (Test-AnnouncementContentsEqual -DriveId $driveId -OldItem $oldItem -NewItem $newItem) {
                    if ($PSCmdlet.ShouldProcess("$folderName/CHU_Y.txt ($sourceUpn)", "Delete duplicate CHU_Y.txt (content identical to .CHU_Y.txt)")) {
                        try {
                            Invoke-GraphApiWithRetry -Uri "https://graph.microsoft.com/v1.0/drives/$driveId/items/$($oldItem.id)" -Method DELETE | Out-Null
                            Write-Log "  [$folderName] .CHU_Y.txt already present; removed duplicate CHU_Y.txt (identical content)" -Type SUCCESS
                        } catch {
                            Write-Log "  [$folderName] .CHU_Y.txt already present; failed removing CHU_Y.txt: $_" -Type WARN
                        }
                    } else {
                        Write-Log "  [$folderName] .CHU_Y.txt already present; duplicate CHU_Y.txt kept (-WhatIf)." -Type INFO
                    }
                } else {
                    Write-Log "  [$folderName] .CHU_Y.txt and CHU_Y.txt have DIFFERENT contents - keeping both files." -Type WARN
                }
            } else {
                Write-Log "  [$folderName] Already has .CHU_Y.txt" -Type INFO
            }
            continue
        }
        
        # Case 2: Only CHU_Y.txt exists -> rename to .CHU_Y.txt
        if ($oldItem) {
            $patchBody = @{ name = ".CHU_Y.txt" } | ConvertTo-Json
            try {
                if ($PSCmdlet.ShouldProcess("$folderName/CHU_Y.txt ($sourceUpn)", "Rename to .CHU_Y.txt")) {
                    $renamed = Invoke-GraphApiWithRetry -Uri "https://graph.microsoft.com/v1.0/drives/$driveId/items/$($oldItem.id)" -Method PATCH -Body $patchBody
                    Write-Log "  [$folderName] Renamed: CHU_Y.txt -> .CHU_Y.txt" -Type SUCCESS
                    $renamedCount++
                } else {
                    Write-Log "  [$folderName] Rename skipped (-WhatIf)." -Type INFO
                }
            } catch {
                Write-Log "  [$folderName] Failed to rename CHU_Y.txt: $_" -Type ERROR
            }
            continue
        }
        
        # Case 3: Neither exists in this folder (folder might be empty or never had announcement)
        # We do nothing here
    }
    
    $userStatus = if ($renamedCount -gt 0) { "Renamed ($renamedCount files)" } elseif ($alreadyCount -gt 0) { "Already Done ($alreadyCount files)" } else { "No Announcement Found" }
    Write-Log "Result for ${sourceUpn}: $userStatus" -Type $(if ($renamedCount -gt 0 -or $alreadyCount -gt 0) { 'SUCCESS' } else { 'WARN' })
    
    $summary += [PSCustomObject]@{
        SourceUPN      = $sourceUpn
        RenamedCount   = $renamedCount
        AlreadyRenamed = $alreadyCount
        Status         = $userStatus
    }
}

Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host "   RENAME ANNOUNCEMENT SUMMARY" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
$summary | Format-Table -AutoSize | Out-String | Write-Host -ForegroundColor Cyan

$totalRenamed = ($summary | Measure-Object -Property RenamedCount -Sum).Sum
$totalAlready = ($summary | Measure-Object -Property AlreadyRenamed -Sum).Sum
Write-Log "ALL DONE! Renamed: $totalRenamed | Already up-to-date: $totalAlready" -Type SUCCESS
