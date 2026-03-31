# =============================================================================
# avatar-entraid.ps1
# Azure AI Speech – Text-to-Speech Avatar video using Microsoft Entra ID auth
# =============================================================================
# Prerequisites:
#   - Az.Accounts PowerShell module: Install-Module Az.Accounts -AllowClobber
#   - Logged in: Connect-AzAccount
#   - Custom domain enabled on the Speech resource
#   - RBAC role "Cognitive Services Speech User" assigned to your account
#
# Usage:
#   1. Set $resourceName and avatar settings below
#   2. Place SSML XML files in the .\ssml\ folder
#   3. Run: .\avatar-entraid.ps1
#   4. MP4 video files will be created in the .\video\ folder
#
# Avatar types:
#   Standard video avatar – set $avatarCharacter (e.g. "lisa") and $avatarStyle (e.g. "graceful-sitting")
#   Standard photo avatar – set $avatarCharacter (e.g. "Amara") and leave $avatarStyle empty
#   Custom photo avatar   – set $avatarCharacter to your custom model name,
#                           leave $avatarStyle empty, and set $customAvatar = $true
# =============================================================================

Import-Module Az.Accounts

# ---------- Configuration ----------
$resourceName = "<YOUR-SPEECH-RESOURCE-NAME>"   # e.g. "my-speech-resource"

# Avatar settings
$avatarCharacter = "lisa"             # Avatar character name (see README for full list)
$avatarStyle     = "graceful-sitting" # Avatar style (leave empty "" for photo avatars)
$customAvatar    = $false             # Set to $true for custom photo avatar
$videoFormat     = "Mp4"              # "Mp4" or "webm" (webm needed for transparent background)
$videoCodec      = "h264"             # "h264", "hevc", "vp9", "av1"
$subtitleType    = "soft_embedded"    # "soft_embedded", "hard_embedded", "external_file", "none"
$backgroundColor = "#FFFFFFFF"        # "#RRGGBBAA" – white opaque by default

$inDir  = ".\ssml"
$outDir = ".\video"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$apiBase = "https://$resourceName.cognitiveservices.azure.com"
$apiVersion = "2024-08-01"
$pollIntervalSec = 10

# ---------- Acquire Entra ID token ----------
$tokenResult = Get-AzAccessToken -ResourceUrl "https://cognitiveservices.azure.com"

# Az.Accounts v3+ returns Token as SecureString — convert to plain text
if ($tokenResult.Token -is [System.Security.SecureString]) {
    $accessToken = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR(
        [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenResult.Token))
} else {
    $accessToken = $tokenResult.Token
}

# ---------- Process SSML files ----------
$files = @(Get-ChildItem $inDir -Filter *.xml)
$total = $files.Count
$overwriteAll = $false
$skipAll = $false

for ($i = 0; $i -lt $total; $i++) {
    $file = $files[$i]
    $ssml = Get-Content $file.FullName -Raw
    $extension = if ($videoFormat -eq "webm") { "webm" } else { "mp4" }
    $outFile = Join-Path $outDir ($file.BaseName + ".$extension")

    Write-Progress -Activity "Generating avatar videos" -Status "$($file.Name) ($($i+1)/$total)" `
        -PercentComplete (($i / $total) * 100)

    # Check for existing file
    if (Test-Path $outFile) {
        if ($skipAll) { Write-Host "Skipped: $($file.BaseName).$extension"; continue }
        if (-not $overwriteAll) {
            $choice = Read-Host "$($file.BaseName).$extension already exists. [S]kip / [O]verwrite / Skip [A]ll / Overwrite A[l]l"
            switch ($choice.ToUpper()) {
                'S' { Write-Host "Skipped: $($file.BaseName).$extension"; continue }
                'A' { $skipAll = $true; Write-Host "Skipped: $($file.BaseName).$extension"; continue }
                'L' { $overwriteAll = $true }
            }
        }
    }

    # Build avatar config
    $avatarConfig = @{
        talkingAvatarCharacter = $avatarCharacter
        videoFormat            = $videoFormat
        videoCodec             = $videoCodec
        subtitleType           = $subtitleType
        backgroundColor        = $backgroundColor
        bitrateKbps            = 2000
        customized             = $customAvatar
    }
    if ($avatarStyle -ne "") {
        $avatarConfig["talkingAvatarStyle"] = $avatarStyle
    }

    $body = @{
        inputKind    = "SSML"
        inputs       = @(@{ content = $ssml })
        avatarConfig = $avatarConfig
    } | ConvertTo-Json -Depth 5

    # Use file base name + short GUID as unique synthesis ID
    $synthesisId = "$($file.BaseName)-$([guid]::NewGuid().ToString('N').Substring(0,8))"
    $uri = "$apiBase/avatar/batchsyntheses/${synthesisId}?api-version=$apiVersion"

    $headers = @{
        "Authorization" = "Bearer $accessToken"
        "Content-Type"  = "application/json"
    }

    # Submit batch synthesis job
    Write-Host "Submitting avatar job for $($file.Name) (ID: $synthesisId)..."
    try {
        Invoke-RestMethod -Method Put -Uri $uri -Headers $headers -Body $body | Out-Null
    }
    catch {
        Write-Host "ERROR submitting $($file.Name): $($_.Exception.Message)" -ForegroundColor Red
        continue
    }

    # Poll for completion
    $getHeaders = @{ "Authorization" = "Bearer $accessToken" }
    $status = "NotStarted"
    while ($status -notin @("Succeeded", "Failed")) {
        Start-Sleep -Seconds $pollIntervalSec
        $result = Invoke-RestMethod -Method Get -Uri $uri -Headers $getHeaders
        $status = $result.status
        Write-Host "  [$($file.BaseName)] Status: $status"
    }

    if ($status -eq "Succeeded") {
        # Download the video
        $videoUrl = $result.outputs.result
        Invoke-RestMethod -Method Get -Uri $videoUrl -OutFile $outFile
        Write-Host "Created: $outFile" -ForegroundColor Green
    }
    else {
        Write-Host "FAILED: $($file.Name) – check the Azure portal for details." -ForegroundColor Red
    }
}

Write-Progress -Activity "Generating avatar videos" -Completed
Write-Host "Done: $total files processed."
