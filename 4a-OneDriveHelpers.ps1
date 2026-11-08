<#
.SYNOPSIS
Helper functions for OneDrive migration using Microsoft Graph PowerShell SDK.

Requires (minimum Graph scopes): User.Read.All, Files.Read.All,
Files.ReadWrite.All, Sites.Read.All.
#>

[CmdletBinding()]
param()

# Shared safety helpers (tenant guard, app-only token, required config values).
. "$PSScriptRoot\MigrationHelpers.ps1"

try {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
    [System.Net.ServicePointManager]::DefaultConnectionLimit = 500
    [System.Net.ServicePointManager]::Expect100Continue = $false
    [System.Net.ServicePointManager]::UseNagleAlgorithm = $false
} catch {}

# NOTE (S11): the TimedWebClient class was dead code (defined here, used nowhere
# in the repo) and has been removed.
function Get-UserOneDriveId {
    param ([string]$UserId)
    try {
        $response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$UserId/drives" -ErrorAction Stop
    } catch {
        return $null
    }
    $drives = $response.value
    $userDrive = $drives | Where-Object { $_.name -eq 'OneDrive' } | Select-Object -First 1
    if (-not $userDrive) { $userDrive = $drives | Where-Object { $_.driveType -eq 'business' -and $_.description -notmatch 'Cache' } | Select-Object -First 1 }
    # S11: no first-drive fallback - picking the first drive can select the wrong drive.
    if (-not $userDrive) { throw "Could not identify a unique OneDrive drive for user '$UserId' (found $(@($drives).Count) drive(s), none named 'OneDrive' or business non-cache)." }
    return $userDrive.id
}

function Sanitize-LocalPath([string]$path) {
    if (-not $path) { return $path }
    $prefix = ""
    $main = $path
    if ($path -match '^\\\\\?\\([a-zA-Z]:\\)(.*)$') {
        $prefix = "\\?\" + $Matches[1]
        $main = $Matches[2]
    } elseif ($path -match '^([a-zA-Z]:\\)(.*)$') {
        $prefix = $Matches[1]
        $main = $Matches[2]
    }
    $parts = $main -split '[\\/]'
    for ($i = 0; $i -lt $parts.Count; $i++) {
        $parts[$i] = $parts[$i] -replace '[?*:""<>|]', '_'
        $parts[$i] = $parts[$i].TrimEnd('. ')
    }
    return $prefix + ($parts -join '\')
}

function Get-DriveItemsRecursive {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [string]$DriveItemId = 'root',
        
        [string]$BasePath = '',
        
        [string]$DriveId = ''
    )
    
    $items = [System.Collections.Generic.List[PSCustomObject]]::new()
    try {
        if (-not $DriveId) {
            $DriveId = Get-UserOneDriveId -UserId $UserId
        }
        $displayPath = if ($BasePath) { $BasePath } else { "root" }
        Write-Host "   -> Scanning folder: /$displayPath ..." -ForegroundColor DarkGray
        $children = Get-MgUserDriveItemChild -UserId $UserId -DriveId $DriveId -DriveItemId $DriveItemId -All -ErrorAction Stop
        
        foreach ($child in $children) {
            $isFolder = ($null -ne $child.Folder.ChildCount) -and ($child.Name -notmatch '\.(one|onetoc2)$') -and ($child.Package.Type -ne 'oneNote')
            $itemPath = if ($BasePath) { "$BasePath/$($child.Name)" } else { $child.Name }
            
            $itemObj = [PSCustomObject]@{
                Id           = $child.Id
                Name         = $child.Name
                Path         = $itemPath
                Size         = $child.Size
                IsFolder     = $isFolder
                ParentPath   = $BasePath
                MimeType     = if ($child.File) { $child.File.MimeType } else { $null }
                LastModified = $child.LastModifiedDateTime
            }
            
            $items.Add($itemObj)
            
            if ($isFolder) {
                try {
                    $items.AddRange([PSCustomObject[]]@(Get-DriveItemsRecursive -UserId $UserId -DriveId $DriveId -DriveItemId $child.Id -BasePath $itemPath))
                } catch {
                    Write-Host " [WARN] Could not list children for '$itemPath' (may be a file or package): $_" -ForegroundColor DarkYellow
                }
            }
        }
    } catch {
        Write-Host " [WARN] Error getting items for user ${UserId}, drive item ${DriveItemId}: $_" -ForegroundColor Yellow
    }
    
    return $items
}

function Download-FileStreamWithTimeout {
    param(
        [Parameter(Mandatory=$true)][string]$Url,
        [Parameter(Mandatory=$true)][string]$DestinationPath,
        [int]$ConnectTimeoutSec = 30,
        [int]$ReadTimeoutSec = 30
    )
    $req = [System.Net.HttpWebRequest]::Create($Url)
    $req.Timeout = $ConnectTimeoutSec * 1000
    $req.ReadWriteTimeout = $ReadTimeoutSec * 1000
    $req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    $req.AllowAutoRedirect = $true
    
    $resp = $req.GetResponse()
    try {
        $stream = $resp.GetResponseStream()
        $stream.ReadTimeout = $ReadTimeoutSec * 1000
        $fs = [System.IO.File]::Create($DestinationPath)
        try {
            $buffer = New-Object byte[] 65536
            while (($read = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                $fs.Write($buffer, 0, $read)
            }
        } finally {
            $fs.Close()
            $fs.Dispose()
            $stream.Close()
            $stream.Dispose()
        }
    } finally {
        $resp.Close()
    }
}

function Download-DriveItem {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $true)]
        [string]$DriveItemId,
        
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,
        
        [string]$DriveId = '',
        
        [string]$DownloadUrl = '',
        
        [int]$MaxRetries = 3
    )
    
    $attempt = 0
    $success = $false
    $permanentlyFailed = $false
    $safeDestPath = $DestinationPath  # sanitized actual path on disk (see below)

    if (-not $DriveId) {
        try {
            $userDrive = @(Get-MgUserDrive -UserId $UserId -ErrorAction Stop)[0]
            $DriveId = $userDrive.Id
        } catch {
            Write-Warning "Could not get default DriveId for user ${UserId}: $_"
        }
    }
    
    while ($attempt -lt $MaxRetries -and -not $success -and -not $permanentlyFailed) {
        $attempt++
        try {
            # Sanitize invalid Windows filesystem characters and bypass 260 character MAX_PATH limit
            $safeDestPath = Sanitize-LocalPath $DestinationPath
            if ($safeDestPath -match '^[a-zA-Z]:\\' -and $safeDestPath -notmatch '^\\\\\?\\') {
                $safeDestPath = "\\?\" + $safeDestPath
            }
            
            # Ensure parent directory exists
            $parentDir = Split-Path -Path $safeDestPath -Parent
            if (-not (Test-Path -LiteralPath $parentDir)) {
                [System.IO.Directory]::CreateDirectory($parentDir) | Out-Null
            }
            if (Test-Path -LiteralPath $safeDestPath) {
                Remove-Item -LiteralPath $safeDestPath -Force -ErrorAction SilentlyContinue
            }
            
            $urlToUse = $DownloadUrl
            if (-not $urlToUse) {
                try {
                    $itemObj = Get-MgUserDriveItem -UserId $UserId -DriveId $DriveId -DriveItemId $DriveItemId -ErrorAction Stop
                    $urlToUse = $itemObj.AdditionalProperties['@microsoft.graph.downloadUrl']
                } catch {}
            }
            $downloaded = $false
            if ($urlToUse) {
                try {
                    Download-FileStreamWithTimeout -Url $urlToUse -DestinationPath $safeDestPath -ConnectTimeoutSec 30 -ReadTimeoutSec 30
                    $downloaded = $true
                } catch {
                    Write-Verbose "CDN stream download failed, falling back to Graph API: $_"
                }
            }
            if (-not $downloaded) {
                Get-MgUserDriveItemContent -UserId $UserId -DriveId $DriveId -DriveItemId $DriveItemId -OutFile $safeDestPath -ErrorAction Stop
            }
            $success = $true
        } catch {
            $errMsg = $_.Exception.Message
            if (-not $errMsg) { $errMsg = "$_" }
            if ($errMsg.Length -gt 150) { $errMsg = $errMsg.Substring(0, 150) + "..." }
            
            # Detect permanent failures — no point retrying these.
            # S11: a 403 is only permanent when there is NO Retry-After header;
            # with Retry-After it is throttling and must be retried.
            $hasRetryAfter = $false
            try {
                $errResponse = $_.Exception.Response
                if ($null -ne $errResponse -and $null -ne $errResponse.Headers -and $errResponse.Headers['Retry-After']) { $hasRetryAfter = $true }
            } catch {}
            if (-not $hasRetryAfter -and $errMsg -match 'Retry-After') { $hasRetryAfter = $true }
            $is403 = ($errMsg -match '(403|Access Denied|Forbidden)')
            $isPermanent = ($errMsg -match '(404|Not Found|Unable to find the specified file|forbidden by antivirus|499|itemNotFound|could not be found|451|Unavailable)') -or ($is403 -and -not $hasRetryAfter)
            if ($isPermanent) {
                Write-Warning "Permanently skipping item $DriveItemId : $errMsg"
                $permanentlyFailed = $true
            } else {
                Write-Warning "Attempt $attempt failed to download item $DriveItemId for user $UserId : $errMsg"
                if (Test-Path -LiteralPath $safeDestPath) {
                    Remove-Item -LiteralPath $safeDestPath -Force -ErrorAction SilentlyContinue
                }
                if ($attempt -lt $MaxRetries) {
                    $sleepTime = [math]::Pow(2, $attempt)
                    Start-Sleep -Seconds $sleepTime
                }
            }
        }
    }
    
    # Return hashtable so caller can distinguish transient vs permanent failure.
    # SavedPath is the sanitized path actually written to disk; callers must check
    # this path (not the raw DestinationPath) for existence afterwards.
    return @{ Success = $success; PermanentlyFailed = $permanentlyFailed; SavedPath = $safeDestPath }

}

function Upload-DriveItem {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $true)]
        [string]$DestDriveId,
        
        [Parameter(Mandatory = $true)]
        [string]$ParentItemId,
        
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        
        [Parameter(Mandatory = $true)]
        [string]$FileName,
        
        [int]$MaxRetries = 3,
        
        [int]$ChunkSize = 10MB
    )
    
    $fileInfo = Get-Item -Path $FilePath
    $fileSize = $fileInfo.Length
    $fourMB = 4MB
    $encodedFileName = [System.Uri]::EscapeDataString($FileName)
    
    $attempt = 0
    $result = $null
    $success = $false
    
    while ($attempt -lt $MaxRetries -and -not $success) {
        $attempt++
        try {
            if ($fileSize -le $fourMB) {
                # Simple upload
                $fileBytes = [System.IO.File]::ReadAllBytes($FilePath)
                if ($DestDriveId) {
                    $uri = "https://graph.microsoft.com/v1.0/drives/${DestDriveId}/items/${ParentItemId}:/${encodedFileName}:/content"
                } else {
                    $uri = "https://graph.microsoft.com/v1.0/users/${UserId}/drive/items/${ParentItemId}:/${encodedFileName}:/content"
                }
                $response = Invoke-MgGraphRequest -Method PUT -Uri $uri -Body $fileBytes -ContentType "application/octet-stream" -ErrorAction Stop
                $result = $response
                $success = $true
            } else {
                # Resumable upload session
                if ($DestDriveId) {
                    $uri = "https://graph.microsoft.com/v1.0/drives/${DestDriveId}/items/${ParentItemId}:/${encodedFileName}:/createUploadSession"
                } else {
                    $uri = "https://graph.microsoft.com/v1.0/users/${UserId}/drive/items/${ParentItemId}:/${encodedFileName}:/createUploadSession"
                }
                $body = @{
                    item = @{
                        "@microsoft.graph.conflictBehavior" = "replace"
                    }
                } | ConvertTo-Json -Depth 5
                
                $uploadSession = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType "application/json" -ErrorAction Stop
                $uploadUrl = $uploadSession.uploadUrl
                
                $fileStream = [System.IO.File]::OpenRead($FilePath)
                $buffer = New-Object byte[] $ChunkSize
                $bytesRead = 0
                $startRange = 0
                
                while (($bytesRead = $fileStream.Read($buffer, 0, $ChunkSize)) -gt 0) {
                    $endRange = $startRange + $bytesRead - 1
                    $contentRange = "bytes $startRange-$endRange/$fileSize"
                    
                    $chunk = $buffer
                    if ($bytesRead -lt $ChunkSize) {
                        $chunk = New-Object byte[] $bytesRead
                        [System.Array]::Copy($buffer, $chunk, $bytesRead)
                    }
                    
                    $headers = @{
                        "Content-Length" = $bytesRead.ToString()
                        "Content-Range"  = $contentRange
                    }
                    
                    $chunkUploaded = $false
                    $chunkAttempt = 0
                    while (-not $chunkUploaded -and $chunkAttempt -lt 5) {
                        $chunkAttempt++
                        try {
                            $uploadResponse = Invoke-WebRequest -Method PUT -Uri $uploadUrl -Body $chunk -Headers $headers -UseBasicParsing -TimeoutSec 3600 -ErrorAction Stop
                            $chunkUploaded = $true
                        } catch {
                            if ($chunkAttempt -lt 5) {
                                Write-Host " [RETRY CHUNK] Network drop detected on chunk ($contentRange), retrying in 5s... ($chunkAttempt/5)" -ForegroundColor Yellow
                                Start-Sleep -Seconds 5
                            } else {
                                throw $_
                            }
                        }
                    }
                    
                    $startRange += $bytesRead
                }
                $result = $uploadResponse
                $success = $true
                
                $fileStream.Close()
            }
        } catch {
            $errMsg = $_.Exception.Message
            if (-not $errMsg) { $errMsg = "$_" }
            if ($errMsg.Length -gt 150) { $errMsg = $errMsg.Substring(0, 150) + "..." }
            Write-Warning "Attempt $attempt failed to upload file ${FileName}: $errMsg"
            if ($attempt -lt $MaxRetries) {
                $sleepTime = [math]::Pow(2, $attempt)
                Start-Sleep -Seconds $sleepTime
            }
            if ($null -ne $fileStream) {
                $fileStream.Close()
            }
        }
    }
    
    return $result
}

function New-DriveFolder {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $false)]
        [string]$DriveId,
        
        [Parameter(Mandatory = $true)]
        [string]$ParentItemId,
        
        [Parameter(Mandatory = $true)]
        [string]$FolderName
    )
    
    # S11: encode the folder name interpolated into the $filter URI (single quotes
    # doubled for OData, then URL-encoded like file names in Upload-DriveItem).
    $escapedFolderName = [System.Uri]::EscapeDataString(($FolderName -replace "'", "''"))
    if ($DriveId) {
        $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ParentItemId/children"
        $existingUri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ParentItemId/children?`$filter=name eq '$escapedFolderName'"
    } else {
        $uri = "https://graph.microsoft.com/v1.0/users/$UserId/drive/items/$ParentItemId/children"
        $existingUri = "https://graph.microsoft.com/v1.0/users/$UserId/drive/items/$ParentItemId/children?`$filter=name eq '$escapedFolderName'"
    }
    $body = @{
        name = $FolderName
        folder = @{}
        "@microsoft.graph.conflictBehavior" = "rename"
    } | ConvertTo-Json -Depth 5
    
    try {
        # Check if folder exists
        $existing = Invoke-MgGraphRequest -Method GET -Uri $existingUri -ErrorAction SilentlyContinue
        if ($existing.value.Count -gt 0) {
            return $existing.value[0]
        }
        
        $folder = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType "application/json" -ErrorAction Stop
        return $folder
    } catch {
        Write-Warning "Failed to create folder ${FolderName} for user ${UserId}: $_"
        return $null
    }
}

function Get-AllDriveItemsFast {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserId
    )
    
    $items = [System.Collections.Generic.List[PSCustomObject]]::new()
    try {
        Write-Host "   -> [Scan] Retrieving OneDrive files via Delta API (paged)..." -ForegroundColor Cyan
        $DriveId = Get-UserOneDriveId -UserId $UserId
        
        $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/root/delta?`$select=id,name,size,folder,package,parentReference,file,lastModifiedDateTime,@microsoft.graph.downloadUrl,deleted"
        $pageCount = 0
        
        while ($uri) {
            $pageCount++
            $response = $null
            for ($attempt = 1; $attempt -le 5; $attempt++) {
                try {
                    $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
                    break
                } catch {
                    if ($attempt -lt 5) {
                        Write-Host "      [RETRY] Delta API page $pageCount failed ($($_.Exception.Message)). Retrying in $($attempt * 3)s ($attempt/5)..." -ForegroundColor Yellow
                        Start-Sleep -Seconds ($attempt * 3)
                    } else {
                        throw $_
                    }
                }
            }
            $deltaItems = $response.value
            
            foreach ($child in $deltaItems) {
                $name = $child.name
                if (-not $name -or $name -eq 'root') { continue }
                if ($child.deleted) { continue }
                
                $isFolder = ($null -ne $child.folder) -and ($name -notmatch '\.(one|onetoc2)$')
                if ($child.package -and $child.package.type -eq 'oneNote') { $isFolder = $false }
                
                $ParentPath = ""
                if ($child.parentReference -and $child.parentReference.path) {
                    $ParentPath = $child.parentReference.path -replace '^.*root:?/?', ''
                    $ParentPath = $ParentPath -replace '^/', '' -replace '/$', ''
                }
                
                $itemPath = if ($ParentPath) { "$ParentPath/$name" } else { $name }
                
                $itemObj = [PSCustomObject]@{
                    Id           = $child.id
                    Name         = $name
                    Path         = $itemPath
                    Size         = $child.size
                    IsFolder     = $isFolder
                    ParentPath   = $ParentPath
                    MimeType     = if ($child.file) { $child.file.mimeType } else { $null }
                    LastModified = $child.lastModifiedDateTime
                    DownloadUrl  = $child.'@microsoft.graph.downloadUrl'
                }
                $items.Add($itemObj)
            }
            
            Write-Host "      Page $pageCount done, $($items.Count) items found so far..." -ForegroundColor DarkGray
            
            # Next page or delta link
            if ($response.'@odata.nextLink') {
                $uri = $response.'@odata.nextLink'
            } else {
                $uri = $null
            }
        }
        if ($items.Count -eq 0) {
            Write-Host "   -> [Scan] 0 items from Delta API, falling back to recursive scan..." -ForegroundColor DarkYellow
            $items = Get-DriveItemsRecursive -UserId $UserId -DriveId "" -DriveItemId "root" -BasePath ""
        }
        Write-Host "   -> [Scan Complete] $($items.Count) total items found." -ForegroundColor Green
    } catch {
        Write-Host " [WARN] Delta scan failed ($($_)). Falling back to standard scan..." -ForegroundColor DarkYellow
        $items = Get-DriveItemsRecursive -UserId $UserId -DriveId "" -DriveItemId "root" -BasePath ""
    }
    return $items
}

