# PowerShell用ユーティリティモジュール
# 使用方法: . "c:\work\PSUtils.ps1"

# ログ出力方法の定義
enum LogOutputMethod {
    Native         # PowerShellのネイティブ出力（Write-Output, Write-Verbose等）
    PSScriptTools  # PSScriptToolsモジュール
    PSFramework    # PSFrameworkモジュール
    Custom         # カスタムロガー
}

# ログレベル定義
enum LogLevel {
    Debug = 0
    Info  = 1
    Warn  = 2
    Error = 3
}

# ログ設定
$script:LogConfig = @{
    Method = [LogOutputMethod]::Native
    Level = [LogLevel]::Info
    IncludeTimestamp = $true
    IncludeLevel = $true
    OutputToFile = $false
    LogFilePath = ""
    Category = ""
}

# ログ出力方法の初期化
function Initialize-Logger {
    param(
        [LogOutputMethod]$Method           = [LogOutputMethod]::Native,
        [LogLevel]       $Level            = [LogLevel]::Info,
        [bool]           $IncludeTimestamp = $true,
        [bool]           $IncludeLevel     = $true,
        [bool]           $OutputToFile     = $false,
        [string]         $LogFilePath      = "",
        [string]         $Category         = ""
    )
    
    $script:LogConfig.Method            = $Method
    $script:LogConfig.Level             = $Level
    $script:LogConfig.IncludeTimestamp  = $IncludeTimestamp
    $script:LogConfig.IncludeLevel      = $IncludeLevel
    $script:LogConfig.OutputToFile      = $OutputToFile
    $script:LogConfig.LogFilePath       = $LogFilePath
    $script:LogConfig.Category          = $Category
    
    # 選択された方法に応じて初期化
    switch ($Method) {
        ([LogOutputMethod]::PSScriptTools) {
            try {
                Import-Module PSScriptTools -ErrorAction Stop
                Write-Output "PSScriptToolsロガーを初期化しました"
            } catch {
                Write-Warning "PSScriptToolsモジュールが見つかりません。ネイティブ出力に切り替えます。"
                $script:LogConfig.Method = [LogOutputMethod]::Native
            }
        }
        ([LogOutputMethod]::PSFramework) {
            try {
                Import-Module PSFramework -ErrorAction Stop
                Write-Output "PSFrameworkロガーを初期化しました"
            } catch {
                Write-Warning "PSFrameworkモジュールが見つかりません。ネイティブ出力に切り替えます。"
                $script:LogConfig.Method = [LogOutputMethod]::Native
            }
        }
        ([LogOutputMethod]::Custom) {
            Write-Output "カスタムロガーを初期化しました"
        }
        default {
            Write-Output "ネイティブ出力ロガーを初期化しました"
        }
    }
}

# 統一されたログ出力関数
function Write-PSLog {
    param(
        [LogLevel]       $Level    = [LogLevel]::Info,
        [string]         $Message,
        [string]         $Category = ""
    )
    
    # カテゴリの設定
    if ($Category -eq "") {
        $Category = $script:LogConfig.Category
    }
    
    # ログレベルチェック
    if ($Level -lt $script:LogConfig.Level) {
        return
    }
    
    # 選択された方法でログ出力
    switch ($script:LogConfig.Method) {
        ([LogOutputMethod]::PSScriptTools) {
            Write-ScriptLog -Message $Message -Level $Level.ToString()
        }
        ([LogOutputMethod]::PSFramework) {
            Write-PSFMessage -Level $Level.ToString() -Message $Message
        }
        ([LogOutputMethod]::Custom) {
            Write-NativeLog -Message $Message -Level $Level -Category $Category
        }
        default {
            Write-NativeLog -Message $Message -Level $Level -Category $Category
        }
    }
}

# ネイティブ出力ログ関数
function Write-NativeLog {
    param(
        [string]         $Message,
        [LogLevel]       $Level,
        [string]         $Category = ""
    )
    
    # ログメッセージ構築
    $logParts = @()
    
    if ($script:LogConfig.IncludeTimestamp) {
        $logParts += Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    
    if ($script:LogConfig.IncludeLevel) {
        # ログレベル全体を8桁で右パディング
        $levelPart = "[$Level]"
        $levelPart = $levelPart.PadRight(8)
        $logParts += $levelPart
    }
    
    if ($Category -ne "") {
        $logParts += "[$Category]"
    }
    
    $logParts += $Message
    $logMessage = $logParts -join " "
    
    # 出力
    switch ($Level) {
        ([LogLevel]::Debug)   { Write-Verbose $logMessage }
        ([LogLevel]::Info)    { Write-Output  $logMessage }
        ([LogLevel]::Warn)    { Write-Warning $logMessage }
        ([LogLevel]::Error)   { Write-Error   $logMessage }
    }
    
    # ファイル出力
    if ($script:LogConfig.OutputToFile -and $script:LogConfig.LogFilePath -ne "") {
        try {
            $logMessage | Out-File -FilePath $script:LogConfig.LogFilePath -Append -Encoding UTF8
        } catch {
            Write-Error "ログファイルへの書き込みに失敗しました: $($_.Exception.Message)"
        }
    }
}

# ログ設定の変更
function Configure-LogConfig {
    param(
        [LogLevel]       $Level            = [LogLevel]::Info,
        [bool]           $IncludeTimestamp = $true,
        [bool]           $IncludeLevel     = $true,
        [bool]           $OutputToFile     = $false,
        [string]         $LogFilePath      = "",
        [string]         $Category         = ""
    )
    
    $script:LogConfig.Level = $Level
    $script:LogConfig.IncludeTimestamp = $IncludeTimestamp
    $script:LogConfig.IncludeLevel = $IncludeLevel
    $script:LogConfig.OutputToFile = $OutputToFile
    $script:LogConfig.LogFilePath = $LogFilePath
    $script:LogConfig.Category = $Category
}

# ログ出力方法の変更
function Switch-LogMethod {
    param(
        [LogOutputMethod]$Method
    )
    
    Initialize-Logger -Method $Method -Level $script:LogConfig.Level
}

# フォルダ内の最新ファイルを取得する関数
# param1 対象フォルダパス
# param2 ファイル拡張子フィルタ
# 戻り値: 最新ファイルの名前（見つからない場合はnull）
function Get-LatestFile {
    param(
        [string]$folderPath,
        [string]$fileExtension
    )
    
    try {
        # フォルダ内のファイルを更新日時順で取得
        $latestFile = Get-ChildItem -Path $folderPath -Filter $fileExtension | 
                     Sort-Object LastWriteTime -Descending | 
                     Select-Object -First 1
        
        if ($null -ne $latestFile) {
            Write-PSLog -Level ([LogLevel]::Info) -Message "最新ファイルをデフォルトに設定: $($latestFile.Name) (更新日時: $($latestFile.LastWriteTime))"
            return $latestFile.Name
        } else {
            Write-PSLog -Level ([LogLevel]::Warn) -Message "フォルダ内に$fileExtensionファイルが見つかりません"
            return $null
        }
    } catch {
        Write-PSLog -Level ([LogLevel]::Warn) -Message "最新ファイルの取得に失敗しました: $($_.Exception.Message)"
        return $null
    }
}

# 初期化（デフォルトはネイティブ出力）
Initialize-Logger -Method ([LogOutputMethod]::Native) -Level ([LogLevel]::Info) 