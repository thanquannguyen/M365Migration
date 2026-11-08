<#
.SYNOPSIS
    Master orchestrator script for Microsoft 365 Migration.

.DESCRIPTION
    Runs the migration pipeline in script-number order:
      0 Prerequisites, 1 Export users, 2 Create users, 3 Assign licenses,
      3a Provision OneDrive sites, 3b Grant destination OneDrive admin access,
      3c Initialize OneDrive libraries, 4 Migrate OneDrive data (or 6, the
      rclone engine, with -UseRcloneEngine), 4c Clean destination duplicates,
      5 Validate, 5a Announcements, 5b Full source purge, 5c Partial source
      purge, 5d Empty recycle bin, 5e Rename announcement.

    -StartFromPhase / -EndAtPhase take a phase NUMBER (0-5) for compatibility.
    A numeric bound covers its whole phase group: -EndAtPhase 3 runs through
    3c, -EndAtPhase 4 runs through 4c, -EndAtPhase 5 runs through 5e.
    -StartFromPhase 4 starts at phase 4 (3a/3b/3c are skipped).

    Destructive phases (5b, 5c, 5d) ask for confirmation BEFORE running, and
    only then pass -Confirm to the script; on their own the purge scripts
    prompt for confirmation before deleting (ConfirmImpact High), and
    -WhatIf previews without deleting. -PermanentDelete is never passed
    by this orchestrator: permanent delete is a separate, manual step
    (see README). The purge scripts additionally enforce the validation gate
    themselves: they refuse to run unless a recent passing 5-Validate.ps1
    result exists.

.PARAMETER UseRcloneEngine
    Run phase 4 with 6-RcloneMigration.ps1 (rclone engine) instead of
    4-MigrateOneDrive.ps1 (Graph engine, the default).
#>
[CmdletBinding()]
param (
    [string]$ConfigPath = '.\MigrationConfig.psd1',
    [int]$StartFromPhase = 0,
    [int]$EndAtPhase = 5,
    [switch]$SkipOneDrive,
    [switch]$UseRcloneEngine
)

# Shared safety helpers.
. "$PSScriptRoot\MigrationHelpers.ps1"

$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"
$env:AZURE_IDENTITY_DISABLE_WAM = "true"

$StartTime = Get-Date
$LogDir = ".\Logs"
$OutputDir = ".\Output"
$StagingDir = ".\MigrationStaging"

# Pipeline order. Letter phases run immediately after their number phase.
$PhaseOrder = @('0', '1', '2', '3', '3a', '3b', '3c', '4', '4c', '5', '5a', '5b', '5c', '5d', '5e')

function Get-PhaseIndex {
    param([string]$PhaseId)
    return [Array]::IndexOf($PhaseOrder, $PhaseId)
}

function Get-PhaseGroupEndIndex {
    # Last phase index belonging to numeric phase group $Number
    # (e.g. 3 -> 3c, 4 -> 4c, 5 -> 5e).
    param([int]$Number)
    $end = -1
    for ($i = 0; $i -lt $PhaseOrder.Count; $i++) {
        if ($PhaseOrder[$i] -match "^$Number[a-z]?$") { $end = $i }
    }
    return $end
}

$StartPhaseIndex = Get-PhaseIndex -PhaseId "$StartFromPhase"
if ($StartPhaseIndex -lt 0) {
    throw "Invalid -StartFromPhase: $StartFromPhase. Valid phase numbers are 0-5."
}
$EndPhaseIndex = Get-PhaseGroupEndIndex -Number $EndAtPhase
if ($EndPhaseIndex -lt 0) {
    throw "Invalid -EndAtPhase: $EndAtPhase. Valid phase numbers are 0-5."
}

# Create directories
foreach ($Dir in @($LogDir, $OutputDir, $StagingDir)) {
    if (-not (Test-Path $Dir)) {
        New-Item -ItemType Directory -Path $Dir -Force | Out-Null
    }
}

$LogFile = Join-Path $LogDir "Migration_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param([string]$Message, [string]$Color = "Cyan")
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] $Message"
    Add-Content -Path $LogFile -Value $LogMessage
    Write-Host $LogMessage -ForegroundColor $Color
}

function Prompt-Continue {
    param([string]$Message)
    while ($true) {
        $Response = Read-Host "$Message"
        if ($Response -match '^[Yy]$') { return $true }
        if ($Response -match '^[Nn]$') { return $false }
    }
}

function Run-Phase {
    param(
        [string]$PhaseId,
        [string]$ScriptName,
        [string]$PromptMsg,
        [scriptblock]$PostAction = $null,
        [string]$PrePrompt = $null,
        [hashtable]$ExtraArgs = @{}
    )

    $phaseIndex = Get-PhaseIndex -PhaseId $PhaseId
    if ($phaseIndex -lt 0) {
        Write-Log "Unknown Phase $PhaseId ($ScriptName); skipping." "Red"
        return $false
    }
    if ($StartPhaseIndex -gt $phaseIndex) {
        Write-Log "Skipping Phase $PhaseId ($ScriptName) based on -StartFromPhase." "DarkGray"
        return $true
    }
    if ($EndPhaseIndex -lt $phaseIndex) {
        Write-Log "Stopping before Phase $PhaseId ($ScriptName) based on -EndAtPhase." "DarkGray"
        return $false
    }

    Write-Log "`n--- Starting Phase ${PhaseId}: ${ScriptName} ---" "Magenta"

    if ($PrePrompt) {
        $confirmed = Prompt-Continue $PrePrompt
        if (-not $confirmed) {
            Write-Log "User declined Phase $PhaseId ($ScriptName) before it ran. Aborting migration." "Yellow"
            return $false
        }
    }

    $ScriptPath = Join-Path $PSScriptRoot $ScriptName
    if (-not (Test-Path $ScriptPath)) {
        Write-Log "Script $ScriptPath not found!" "Red"
        $Continue = Prompt-Continue "Error running phase $PhaseId. Ignore and continue to next phase? (Y/N)"
        return $Continue
    }

    try {
        & $ScriptPath -ConfigPath $ConfigPath @ExtraArgs
        if ($PostAction) {
            & $PostAction
        }
    } catch {
        Write-Log "Error running ${ScriptName}: $_" "Red"
        $Continue = Prompt-Continue "Error occurred. Continue to next phase? (Y/N)"
        return $Continue
    }

    if ($PromptMsg) {
        $Continue = Prompt-Continue $PromptMsg
        if (-not $Continue) {
            Write-Log "User aborted the migration after Phase $PhaseId." "Yellow"
            return $false
        }
    }

    return $true
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "    Microsoft 365 Migration Tool        " -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

Write-Log "Starting Migration Orchestrator" "Green"
if (Test-Path $ConfigPath) {
    Write-Log "Loading configuration from $ConfigPath" "Yellow"
    try {
        $Config = Import-LocalizedData -BaseDirectory (Split-Path $ConfigPath) -FileName (Split-Path $ConfigPath -Leaf)
        # Never print config values: client secrets and the default password
        # live in this file. Only list the key names that were loaded.
        $loadedKeys = @()
        if ($Config -is [System.Collections.IDictionary]) {
            $loadedKeys = @($Config.Keys | Sort-Object)
        }
        Write-Host "Loaded configuration keys: $($loadedKeys -join ', ')" -ForegroundColor DarkCyan
    } catch {
        Write-Log "Could not parse config file, but proceeding." "DarkGray"
    }
} else {
    Write-Log "Warning: $ConfigPath not found. Scripts may fail." "Red"
}

# Phase 0
$Phase0Result = Run-Phase -PhaseId 0 -ScriptName "0-Prerequisites.ps1" -PromptMsg "Prerequisites verified. Continue to export users? (Y/N)"
if (-not $Phase0Result) { exit }

# Phase 1
$Phase1Action = {
    $ExportPath = ".\Output\ExportedUsers.csv"
    if (Test-Path $ExportPath) {
        $Count = (Import-Csv $ExportPath).Count
        Write-Log "Exported user count: $Count" "Green"
    }
}
$Phase1Result = Run-Phase -PhaseId 1 -ScriptName "1-ExportUsers.ps1" -PromptMsg "Export complete. Continue to create users in destination? (Y/N)" -PostAction $Phase1Action
if (-not $Phase1Result) { exit }

# Phase 2
$Phase2Action = {
    $MappingPath = ".\Output\UserMapping.csv"
    if (Test-Path $MappingPath) {
        $Count = (Import-Csv $MappingPath).Count
        Write-Log "Created user count: $Count" "Green"
    }
}
$Phase2Result = Run-Phase -PhaseId 2 -ScriptName "2-CreateUsers.ps1" -PromptMsg "Users created. Continue to assign licenses? (Y/N)" -PostAction $Phase2Action
if (-not $Phase2Result) { exit }

# Phase 3
$Phase3Prompt = if ($SkipOneDrive) { "Licenses assigned. OneDrive phases are skipped. Run validation? (Y/N)" } else { "Licenses assigned. Continue to provision OneDrive sites? (Y/N)" }
$Phase3Action = {
    Write-Log "License assignment phase completed." "Green"
}
$Phase3Result = Run-Phase -PhaseId 3 -ScriptName "3-AssignLicenses.ps1" -PromptMsg $Phase3Prompt -PostAction $Phase3Action
if (-not $Phase3Result) { exit }

# Phase 3a / 3b / 3c — OneDrive provisioning
if (-not $SkipOneDrive) {
    $Phase3aResult = Run-Phase -PhaseId '3a' -ScriptName "3a-ProvisionOneDriveSites.ps1" -PromptMsg "OneDrive sites provisioned. Continue to grant admin access to destination sites? (Y/N)"
    if (-not $Phase3aResult) { exit }

    $Phase3bResult = Run-Phase -PhaseId '3b' -ScriptName "3b-GrantDestOneDriveAdminAccess.ps1" -PromptMsg "Admin access granted. Continue to initialize OneDrive libraries? (Y/N)"
    if (-not $Phase3bResult) { exit }

    $Phase3cResult = Run-Phase -PhaseId '3c' -ScriptName "3c-InitOneDriveLibraries.ps1" -PromptMsg "Libraries initialized. Continue to migrate OneDrive data? (Y/N)"
    if (-not $Phase3cResult) { exit }
} else {
    Write-Log "Skipping Phases 3a/3b/3c (OneDrive provisioning) due to -SkipOneDrive switch." "Yellow"
}

# Phase 4 — data migration (Graph engine by default, rclone engine with -UseRcloneEngine)
if (-not $SkipOneDrive) {
    $Phase4Script = if ($UseRcloneEngine) { "6-RcloneMigration.ps1" } else { "4-MigrateOneDrive.ps1" }
    $Phase4Engine = if ($UseRcloneEngine) { "rclone (6-RcloneMigration.ps1)" } else { "Graph (4-MigrateOneDrive.ps1)" }
    Write-Log "Phase 4 engine: $Phase4Engine" "Cyan"
    $Phase4Action = {
        Write-Log "OneDrive migration phase completed." "Green"
    }
    $Phase4Result = Run-Phase -PhaseId 4 -ScriptName $Phase4Script -PromptMsg "OneDrive migration complete. Clean destination duplicates? (Y/N)" -PostAction $Phase4Action
    if (-not $Phase4Result) { exit }

    # Phase 4c
    $Phase4cResult = Run-Phase -PhaseId '4c' -ScriptName "4c-CleanDestDuplicates.ps1" -PromptMsg "Duplicate cleanup complete. Run validation? (Y/N)"
    if (-not $Phase4cResult) { exit }
} else {
    Write-Log "Skipping Phase 4/4c (OneDrive Migration) due to -SkipOneDrive switch." "Yellow"
}

# Phase 5
$Phase5Result = Run-Phase -PhaseId 5 -ScriptName "5-Validate.ps1" -PromptMsg "Validation complete. Continue to post announcements? (Y/N)"
if (-not $Phase5Result) { exit }

# Phase 5a — announcements
$Phase5aResult = Run-Phase -PhaseId '5a' -ScriptName "5a-Announcement.ps1" -PromptMsg "Announcements posted. Continue to purge source data? (Y/N)"
if (-not $Phase5aResult) { exit }

# Phase 5b — DESTRUCTIVE: full source purge (files go to the recycle bin).
$Phase5bResult = Run-Phase -PhaseId '5b' -ScriptName "5b-PurgeSource.ps1" `
    -PrePrompt "DESTRUCTIVE PHASE 5b: this deletes ALL files on source OneDrives (except announcements; deleted files go to the recycle bin, recoverable for 93 days). The validation gate must have passed. Really run the full source purge now? (Y/N)" `
    -ExtraArgs @{ Confirm = $true } `
    -PromptMsg "Full source purge complete. Continue to partial purge? (Y/N)"
if (-not $Phase5bResult) { exit }

# Phase 5c — DESTRUCTIVE: partial source purge (quota-driven; files go to the recycle bin unless -PermanentDelete, which this orchestrator never passes).
$Phase5cResult = Run-Phase -PhaseId '5c' -ScriptName "5c-PartialPurgeSource.ps1" `
    -PrePrompt "DESTRUCTIVE PHASE 5c: this deletes old files on source OneDrives (deleted files go to the recycle bin; permanent delete is NOT used here). The validation gate must have passed. Really run the partial source purge now? (Y/N)" `
    -ExtraArgs @{ Confirm = $true } `
    -PromptMsg "Partial source purge complete. Continue to empty recycle bins? (Y/N)"
if (-not $Phase5cResult) { exit }

# Phase 5d — DESTRUCTIVE: permanently empties OneDrive recycle bins (unrecoverable).
$Phase5dResult = Run-Phase -PhaseId '5d' -ScriptName "5d-EmptyRecycleBin.ps1" `
    -PrePrompt "DESTRUCTIVE PHASE 5d: this PERMANENTLY empties OneDrive recycle bins. Items deleted here CANNOT be recovered. The validation gate must have passed. Really empty the recycle bins now? (Y/N)" `
    -ExtraArgs @{ Confirm = $true } `
    -PromptMsg "Recycle bins emptied. Continue to rename announcements? (Y/N)"
if (-not $Phase5dResult) { exit }

# Phase 5e — rename announcements to .CHU_Y.txt
$Phase5eResult = Run-Phase -PhaseId '5e' -ScriptName "5e-RenameAnnouncement.ps1" -PromptMsg ""
if (-not $Phase5eResult) { exit }

Write-Log "Final report available at .\Output\MigrationReport.html" "Green"

$EndTime = Get-Date
$ElapsedTime = $EndTime - $StartTime
Write-Log "Migration orchestrator finished. Total elapsed time: $($ElapsedTime.ToString('g'))" "Cyan"
