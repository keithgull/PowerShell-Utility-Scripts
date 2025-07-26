

param(
    [string]$sourceFile,
    [string]$targetFolderId
)

# Rcloneアップロード
rclone copy "$sourceFile" "gdrive:$targetFolderId" --drive-shared-with-me

# アップロードファイル名抽出
$filename = [System.IO.Path]::GetFileName($sourceFile)

# Apps Script API URL（例：共有リンク取得）
$apiUrl = "https://script.google.com/macros/s/AKfycbxxxxx/exec?fileName=$filename&folderId=$targetFolderId"

# Web API呼び出し
try {
    $response = Invoke-RestMethod -Uri $apiUrl -Method Get
    $shareUrl = $response.link
} catch {
    $shareUrl = "（共有リンク取得に失敗しました）"
}

# ダイアログ表示
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("共有リンク：" + $shareUrl, "アップロード完了", "OK", "Information")

