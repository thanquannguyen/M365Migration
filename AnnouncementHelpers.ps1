<#
.SYNOPSIS
    Shared announcement template + upload logic for the M365Migration toolkit (S19).

.DESCRIPTION
    Used by 5a-Announcement.ps1 and 5c-PartialPurgeSource.ps1 so the announcement
    text and the 4-folder upload logic exist in exactly one place. Previously
    the template + upload loop were duplicated in both scripts (and both
    embedded the shared default password); fixing one copy left the other
    wrong. The unified template follows 5a (minus the password).

    The template NEVER contains a password (B7): users are told to contact
    their administrator for their password ("contact your administrator for
    your password"). Do not re-add the shared default password here.

    Requires (minimum Graph scopes): Files.ReadWrite.All, Sites.ReadWrite.All
    (delegated, via the caller's Graph session). Callers must dot-source
    4a-OneDriveHelpers.ps1 (Get-UserOneDriveId / New-DriveFolder) before
    calling Publish-AnnouncementToFolders.

    Dot-source from any migration script:

        . "$PSScriptRoot\AnnouncementHelpers.ps1"
#>

function Get-AnnouncementText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceUpn,

        [Parameter(Mandatory = $true)]
        [string]$DestUpn
    )
    # Unified from 5a-Announcement.ps1, with the password sentence (B7)
    # replaced by "contact your administrator for your password".
    $template = [regex]::Unescape("Hi\u1ec7n t\u1ea1i, t\u00e0i kho\u1ea3n cloud l\u01b0u tr\u1eef Onedrive c\u1ee7a b\u1ea1n {0} \u0111\u00e3 chuy\u1ec3n sang {1}. Vui l\u00f2ng li\u00ean h\u1ec7 Qu\u1ea3n tr\u1ecb vi\u00ean (contact your administrator for your password) \u0111\u1ec3 nh\u1eadn m\u1eadt kh\u1ea9u c\u1ee7a b\u1ea1n.`nTo\u00e0n b\u1ed9 d\u1eef li\u1ec7u hi\u1ec7n t\u1ea1i hi\u1ec7n \u0111ang \u0111\u01b0\u1ee3c sao l\u01b0u sang t\u00e0i kho\u1ea3n m\u1edbi v\u00e0 s\u1ebd b\u1ecb xo\u00e1 b\u00ean t\u00e0i kho\u1ea3n n\u00e0y. T\u00e0i kho\u1ea3n hi\u1ec7n t\u1ea1i ch\u1ec9 s\u1eed d\u1ee5ng \u0111\u1ec3 k\u00edch ho\u1ea1t b\u1ea3n quy\u1ec1n Office/Microsoft 365 (Word, Excel, Powerpoint....), h\u00e3y gi\u1eef \u0111\u0103ng nh\u1eadp Office b\u1eb1ng t\u00e0i kho\u1ea3n n\u00e0y.`nH\u00e3y nhanh ch\u00f3ng \u0111\u0103ng nh\u1eadp Onedrive sang t\u00e0i kho\u1ea3n m\u1edbi, n\u1ebfu d\u1eef li\u1ec7u b\u1ea1n c\u1ea7n ch\u01b0a th\u1ea5y sao l\u01b0u qua, h\u00e3y \u0111\u0103ng nh\u1eadp t\u00e0i kho\u1ea3n c\u0169 tr\u00ean web v\u00e0 t\u1ea3i xu\u1ed1ng.`nN\u1ebfu c\u00f3 v\u1ea5n \u0111\u1ec1 v\u1ec1 \u0111\u0103ng nh\u1eadp ho\u1eb7c chuy\u1ec3n t\u00e0i kho\u1ea3n h\u00e3y li\u00ean h\u1ec7 Qu\u1ea3n tr\u1ecb vi\u00ean \u0111\u1ec3 \u0111\u01b0\u1ee3c h\u1ed7 tr\u1ee3. H\u00e3y kh\u1ea9n tr\u01b0\u01a1ng th\u1ef1c hi\u1ec7n \u0111\u1ec3 b\u1ea3o v\u1ec7 d\u1eef li\u1ec7u. To\u00e0n b\u1ed9 d\u1eef li\u1ec7u tr\u00ean t\u00e0i kho\u1ea3n c\u0169 s\u1ebd b\u1ecb xo\u00e1 sau khi qu\u00e1 tr\u00ecnh sao l\u01b0u sang t\u00e0i kho\u1ea3n m\u1edbi c\u1ee7a h\u1ec7 th\u1ed1ng ho\u00e0n t\u1ea5t.")
    return ($template -f $SourceUpn, $DestUpn)
}

function Publish-AnnouncementToFolders {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceUpn,

        [Parameter(Mandatory = $true)]
        [string]$SourceId,

        [Parameter(Mandatory = $true)]
        [string]$DriveId,

        [Parameter(Mandatory = $true)]
        [string]$Content,

        # Optional logging callback: scriptblock taking (Message, Type).
        # Callers pass { param($m, $t) Write-Log $m -Type $t }; without it,
        # progress goes to Write-Host.
        [scriptblock]$LogAction
    )
    $log = {
        param($Message, $Type = 'INFO')
        if ($LogAction) { & $LogAction $Message $Type } else { Write-Host $Message }
    }

    # UTF-8 bytes with BOM, built in memory (zero disk I/O).
    $utf8Bom = New-Object System.Text.UTF8Encoding $true
    $preamble = $utf8Bom.GetPreamble()
    $contentBytes = $utf8Bom.GetBytes($Content)
    $ms = New-Object System.IO.MemoryStream
    $ms.Write($preamble, 0, $preamble.Length)
    $ms.Write($contentBytes, 0, $contentBytes.Length)
    $fileBytes = $ms.ToArray()
    $ms.Dispose()

    $folders = @('root', 'Desktop', 'Documents', 'Pictures')
    $uploadedCount = 0
    & $log "  Announcement generated in-memory ($($fileBytes.Length) bytes), starting uploads..." 'INFO'

    foreach ($folderName in $folders) {
        # S2: retry cap 3 attempts per folder with exponential backoff.
        $uploaded = $false
        for ($attempt = 1; $attempt -le 3 -and -not $uploaded; $attempt++) {
            try {
                if ($folderName -eq 'root') {
                    $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root:/.CHU_Y.txt:/content"
                    Invoke-MgGraphRequest -Method PUT -Uri $uri -Body $fileBytes -ContentType 'application/octet-stream' -ErrorAction Stop | Out-Null
                } else {
                    # New-DriveFolder creates the folder if missing, returns it if present.
                    $folder = New-DriveFolder -UserId $SourceId -DriveId $DriveId -ParentItemId 'root' -FolderName $folderName
                    if (-not $folder) { throw "Could not ensure folder $folderName exists." }
                    $uploadUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$($folder.id):/.CHU_Y.txt:/content"
                    Invoke-MgGraphRequest -Method PUT -Uri $uploadUri -Body $fileBytes -ContentType 'application/octet-stream' -ErrorAction Stop | Out-Null
                }
                $uploaded = $true
                $uploadedCount++
                & $log "  Uploaded .CHU_Y.txt to $folderName for $SourceUpn" 'INFO'
            } catch {
                $err = $_.Exception.Message
                if ($err -match '507|quotaLimitReached|Insufficient Storage') {
                    & $log "  [QUOTA BLOCKED] Cannot upload to $SourceUpn because OneDrive quota limit is reached (Recycle Bin is full). Please empty Recycle Bin on web." 'WARN'
                    return $uploadedCount
                }
                if ($attempt -lt 3) {
                    $delay = 2 * [Math]::Pow(2, $attempt - 1)
                    & $log "  Upload to $folderName for $SourceUpn failed (attempt $attempt/3): $err. Retrying in ${delay}s..." 'WARN'
                    Start-Sleep -Seconds $delay
                } else {
                    & $log "  Failed to upload to $folderName for ${SourceUpn} after 3 attempts: $err" 'ERROR'
                }
            }
        }
    }
    return $uploadedCount
}
