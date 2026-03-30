# 事前: Az PowerShell モジュールが必要 (Install-Module Az.Accounts -AllowClobber)
#       Connect-AzAccount でログイン済みであること
#       Speech リソースにカスタムドメインが設定されていること
#       自分のアカウントに "Cognitive Services Speech User" ロールが付与されていること
Import-Module Az.Accounts

# 変数を設定 — resourceName は Azure ポータルで確認できる Speech リソース名
$resourceName = "aif-legacymodernize"
$inDir  = ".\\ssml"
$outDir = ".\\audio"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

# Entra ID (Azure AD) トークンを取得
$tokenResult = Get-AzAccessToken -ResourceUrl "https://cognitiveservices.azure.com"
# Az.Accounts v3+ は Token を SecureString で返すため変換
if ($tokenResult.Token -is [System.Security.SecureString]) {
  $accessToken = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR(
    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($tokenResult.Token))
} else {
  $accessToken = $tokenResult.Token
}

$files = @(Get-ChildItem $inDir -Filter *.xml)
$total = $files.Count
$overwriteAll = $false
$skipAll = $false

for ($i = 0; $i -lt $total; $i++) {
  $file = $files[$i]
  $ssml = Get-Content $file.FullName -Raw
  $outFile = Join-Path $outDir ($file.BaseName + ".wav")

  Write-Progress -Activity "Generating audio" -Status "$($file.Name) ($($i+1)/$total)" `
    -PercentComplete (($i / $total) * 100)

  # 既存ファイルの確認
  if (Test-Path $outFile) {
    if ($skipAll) { Write-Host "Skipped: $($file.BaseName).wav"; continue }
    if (-not $overwriteAll) {
      $choice = Read-Host "$($file.BaseName).wav already exists. [S]kip / [O]verwrite / Skip [A]ll / Overwrite A[l]l"
      switch ($choice.ToUpper()) {
        'S' { Write-Host "Skipped: $($file.BaseName).wav"; continue }
        'A' { $skipAll = $true; Write-Host "Skipped: $($file.BaseName).wav"; continue }
        'L' { $overwriteAll = $true }
        # 'O' or default: proceed to overwrite this one
      }
    }
  }

  # Entra ID 認証はカスタムドメインエンドポイントが必要
  $uri = "https://$resourceName.cognitiveservices.azure.com/tts/cognitiveservices/v1"
  $headers = @{
    "Authorization"            = "Bearer $accessToken"
    "Content-Type"             = "application/ssml+xml"
    "X-Microsoft-OutputFormat" = "riff-16khz-16bit-mono-pcm"
    "User-Agent"               = "ppt-narration"
  }

  Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $ssml -OutFile $outFile
}
Write-Progress -Activity "Generating audio" -Completed
Write-Host "Done: $total files processed."