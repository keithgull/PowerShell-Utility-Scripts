# ログファイルのパス
$logFile = "C:\Work\log\RestartExplorer.log"

# スクリプト開始時刻を記録
$startTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Add-Content -Path $logFile -Value "Script Start Time: $startTime"


# 開いているエクスプローラウィンドウのフォルダパスを取得してログに記録
#$folders = @()
#$ShellApp = New-Object -ComObject Shell.Application
#foreach ($window in $ShellApp.Windows()) {
#    if ($window.Name -eq "File Explorer") {
#        $folders += $window.Document.Folder.Self.Path
#    }
#}

# 開いているエクスプローラのパスを取得
$explorerPaths = Get-Process | Where-Object { $_.ProcessName -eq "explorer" } | ForEach-Object {
    try {
        [System.Diagnostics.Process]::GetProcessById($_.Id).MainWindowHandle | Out-Null
        ((New-Object -ComObject Shell.Application).Windows() | Where-Object { $_.FullName -like "*explorer.exe" }).Document.Folder.Self.Path
    } catch {
        # エラーが発生した場合は無視
        Write-Host "エラーが発生しました: $_"
    }
}



# 開いているフォルダの数をログに記録
$originalCount = $folders.Count
Add-Content -Path $logFile -Value "Original Explorer Windows Count: $originalCount"

# 各フォルダパスをログに記録
#Add-Content -Path $logFile -Value "Opened Folders:"
#foreach ($folder in $folders) {
#    Add-Content -Path $logFile -Value $folder
#}
# 重複を除外してファイルに出力
$uniquePaths = $explorerPaths | Sort-Object | Get-Unique
Add-Content -Path $logFile -Value $uniquePaths
Write-Host "パスがファイルに出力されました: $logFile"

# エクスプローラプロセスを停止
#Stop-Process -Name "explorer" -Force

# エクスプローラを閉じる
Get-Process explorer | ForEach-Object { Stop-Process -Id $_.Id }


# エクスプローラを再起動
Start-Process "explorer.exe"
Start-Sleep -Seconds 2  # 再起動のため待機

# フォルダを開き直す処理
#$retryCount = 0
#$maxRetries = 5  # 最大リトライ回数
#$openedFolders = @()
#do {
#    Write-Host "Attempt #$($retryCount + 1)"
#    $openedFolders = @()
#    foreach ($folder in $folders) {
#        $process = Start-Process "explorer.exe" -ArgumentList $folder -PassThru
#        if ($process -ne $null) {
#            $openedFolders += $folder
#        }
#    }
#    Start-Sleep -Seconds 2  # 待機
#    $retryCount++
#} while ($openedFolders.Count -ne $originalCount -and $retryCount -lt $maxRetries)

# 新しいエクスプローラを開く
$uniquePaths | ForEach-Object { Start-Process explorer.exe $_ }
Write-Host "新しいエクスプローラが開かれました。"

# 結果をログに記録
if ($openedFolders.Count -eq $originalCount) {
    Add-Content -Path $logFile -Value "All Explorer windows restored successfully."
} else {
    Add-Content -Path $logFile -Value "Failed to restore all Explorer windows."
}

# スクリプト終了時刻を記録
$endTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Add-Content -Path $logFile -Value "Script End Time: $endTime"

# 完了メッセージ
Write-Host "Script completed. Log file saved at $logFile"
