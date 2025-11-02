# PowerShell用ロガーモジュール
# 使用方法: . "c:\work\PSLogger.ps1"

# ログ出力方法の定義
enum LogOutputMethod {
    Native         # PowerShellのネイティブ出力（Write-Output, Write-Verbose等）
    PSScriptTools  # PSScriptToolsモジュール
    PSFramework    # PSFrameworkモジュール
    Custom         # カスタムロガー
    Auto           # ホストの有無を自動判定して出力先を決定
    File           # ファイル出力
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
    LogName = ""
    ScriptPath = ""
}

# 値がnullまたは空文字列かどうかをチェックする関数
# param1 チェック対象の値
# 戻り値: nullまたは空文字列の場合はtrue、それ以外はfalse
function Test-IsNullOrEmpty {
    param(
        [string]$Value
    )
    
    return ($null -eq $Value -or $Value -eq "")
}

# ログ出力方法の初期化
# param1 ログ出力方法
# param2 ログレベル
# param3 タイムスタンプを含めるかどうか
# param4 ログレベルを含めるかどうか
# param5 カテゴリ名
# param6 ログファイルパス（Fileモードの場合）
function Initialize-Logger {
    param(
        [LogOutputMethod]$Method           = [LogOutputMethod]::Native,
        [LogLevel]       $Level            = [LogLevel]::Info,
        [bool]           $IncludeTimestamp = $true,
        [bool]           $IncludeLevel     = $true,
        [string]         $Category         = "",
        [string]         $LogFilePath      = ""
    )
    
    $script:LogConfig.Method            = $Method
    $script:LogConfig.Level             = $Level
    $script:LogConfig.IncludeTimestamp  = $IncludeTimestamp
    $script:LogConfig.IncludeLevel      = $IncludeLevel
    $script:LogConfig.Category          = $Category
    $script:LogConfig.LogName           = ""
    $script:LogConfig.ScriptPath        = ""
    
    # ファイル出力の設定
    if ($Method -eq [LogOutputMethod]::File) {
        $script:LogConfig.OutputToFile = $true
        $script:LogConfig.LogFilePath  = $LogFilePath
        Write-Output "ファイル出力ロガーを初期化しました: $LogFilePath"
    } else {
        $script:LogConfig.OutputToFile = $false
        $script:LogConfig.LogFilePath  = ""
    }
    
    # スクリプトパスを取得（呼び出し元のスクリプトのパス）
    $script:LogConfig.ScriptPath = $MyInvocation.MyCommand.Path
    if (Test-IsNullOrEmpty -Value $script:LogConfig.ScriptPath) {
        $script:LogConfig.ScriptPath = $PSScriptRoot
    }
    if (Test-IsNullOrEmpty -Value $script:LogConfig.ScriptPath) {
        $script:LogConfig.ScriptPath = Get-Location
    }
    
    # 選択された方法に応じて初期化
    switch ($Method) {
        ([LogOutputMethod]::Auto) {
            # ホストの有無を自動判定
            if ([Environment]::UserInteractive -and $Host.Name -ne "ConsoleHost") {
                # ホストがある場合は標準出力
                $script:LogConfig.Method = [LogOutputMethod]::Native
                Write-Output "自動判定: ホスト環境を検出しました。標準出力を使用します。"
            } else {
                # ホストがない場合も標準出力（GDUploadではログファイルは不要）
                $script:LogConfig.Method = [LogOutputMethod]::Native
                Write-Output "自動判定: ホスト環境が検出されませんでした。標準出力を使用します。"
            }
        }
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

# ファイル出力付きログ出力方法の初期化
# param1 ログ出力方法
# param2 ログレベル
# param3 タイムスタンプを含めるかどうか
# param4 ログレベルを含めるかどうか
# param5 カテゴリ名
# param6 ログファイル名（拡張子なし）
# param7 ログファイルパス（指定しない場合はスクリプトのディレクトリに作成）
function Initialize-LoggerWithFile {
    param(
        [LogOutputMethod]$Method           = [LogOutputMethod]::Native,
        [LogLevel]       $Level            = [LogLevel]::Info,
        [bool]           $IncludeTimestamp = $true,
        [bool]           $IncludeLevel     = $true,
        [string]         $Category         = "",
        [string]         $LogName          = "PSLog",
        [string]         $LogFilePath      = ""
    )
    
    $script:LogConfig.Method            = $Method
    $script:LogConfig.Level             = $Level
    $script:LogConfig.IncludeTimestamp  = $IncludeTimestamp
    $script:LogConfig.IncludeLevel      = $IncludeLevel
    $script:LogConfig.OutputToFile      = $true  # ファイル出力を有効化
    $script:LogConfig.Category          = $Category
    $script:LogConfig.LogName           = $LogName
    $script:LogConfig.ScriptPath        = ""
    
    # スクリプトパスを取得（呼び出し元のスクリプトのパス）
    $script:LogConfig.ScriptPath = $MyInvocation.MyCommand.Path
    if (Test-IsNullOrEmpty -Value $script:LogConfig.ScriptPath) {
        $script:LogConfig.ScriptPath = $PSScriptRoot
    }
    if (Test-IsNullOrEmpty -Value $script:LogConfig.ScriptPath) {
        $script:LogConfig.ScriptPath = Get-Location
    }
    
    # ログファイルパスの設定
    if (Test-IsNullOrEmpty -Value $LogFilePath) {
        # ログファイルパスが指定されていない場合はスクリプトのディレクトリに作成
        $scriptPath = Split-Path -Parent $script:LogConfig.ScriptPath
        if (Test-IsNullOrEmpty -Value $scriptPath) {
            $scriptPath = Get-Location
        }
        $script:LogConfig.LogFilePath = Join-Path $scriptPath "$LogName.log"
    } else {
        $script:LogConfig.LogFilePath = $LogFilePath
    }
    
    # ログファイルのディレクトリが存在しない場合は作成
    $logDir = Split-Path -Parent $script:LogConfig.LogFilePath
    if (-not (Test-Path -Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    
    # 選択された方法に応じて初期化
    switch ($Method) {
        ([LogOutputMethod]::Auto) {
            # ホストの有無を自動判定
            if ([Environment]::UserInteractive -and $Host.Name -ne "ConsoleHost") {
                # ホストがある場合は標準出力 + ファイル出力
                $script:LogConfig.Method = [LogOutputMethod]::Native
                Write-Output "自動判定: ホスト環境を検出しました。標準出力 + ファイル出力を使用します: $($script:LogConfig.LogFilePath)"
            } else {
                # ホストがない場合はファイル出力のみ
                $script:LogConfig.Method = [LogOutputMethod]::Custom
                Write-Output "自動判定: ホスト環境が検出されませんでした。ファイル出力のみを使用します: $($script:LogConfig.LogFilePath)"
            }
        }
        ([LogOutputMethod]::PSScriptTools) {
            try {
                Import-Module PSScriptTools -ErrorAction Stop
                Write-Output "PSScriptToolsロガーを初期化しました（ファイル出力: $($script:LogConfig.LogFilePath)）"
            } catch {
                Write-Warning "PSScriptToolsモジュールが見つかりません。ネイティブ出力に切り替えます。"
                $script:LogConfig.Method = [LogOutputMethod]::Native
            }
        }
        ([LogOutputMethod]::PSFramework) {
            try {
                Import-Module PSFramework -ErrorAction Stop
                Write-Output "PSFrameworkロガーを初期化しました（ファイル出力: $($script:LogConfig.LogFilePath)）"
            } catch {
                Write-Warning "PSFrameworkモジュールが見つかりません。ネイティブ出力に切り替えます。"
                $script:LogConfig.Method = [LogOutputMethod]::Native
            }
        }
        ([LogOutputMethod]::Custom) {
            Write-Output "カスタムロガーを初期化しました（ファイル出力: $($script:LogConfig.LogFilePath)）"
        }
        default {
            Write-Output "ネイティブ出力ロガーを初期化しました（ファイル出力: $($script:LogConfig.LogFilePath)）"
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
    if (Test-IsNullOrEmpty -Value $Category) {
        $Category = $script:LogConfig.Category
    }
    
    # ログレベルチェック
    if ($Level -lt $script:LogConfig.Level) {
        return
    }
    
    # 選択された方法でログ出力
    switch ($script:LogConfig.Method) {
        ([LogOutputMethod]::File) {
            Write-NativeLog -Message $Message -Level $Level -Category $Category
        }
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
    
    if (-not (Test-IsNullOrEmpty -Value $Category)) {
        $logParts += "[$Category]"
    }
    
    $logParts += $Message
    $logMessage = $logParts -join " "
    
    # コンソール出力
    switch ($Level) {
        ([LogLevel]::Debug)   { Write-Verbose $logMessage }
        ([LogLevel]::Info)    { Write-Host    $logMessage }
        ([LogLevel]::Warn)    { Write-Warning $logMessage }
        ([LogLevel]::Error)   { Write-Error   $logMessage }
    }
    
    # ファイル出力（設定されている場合のみ）
    if ($script:LogConfig.OutputToFile -and -not (Test-IsNullOrEmpty -Value $script:LogConfig.LogFilePath)) {
        try {
            # ログファイルのディレクトリが存在しない場合は作成
            $logDir = Split-Path -Parent $script:LogConfig.LogFilePath
            if (-not (Test-Path -Path $logDir)) {
                New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            }
            
            $logMessage | Out-File -FilePath $script:LogConfig.LogFilePath -Append -Encoding UTF8
        } catch {
            # エラーをコンソールに出力（ログファイルに書き込めない場合）
            Write-Host "ログファイルへの書き込みに失敗しました: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

# ログ設定の変更
function Configure-LogConfig {
    param(
        [LogLevel]       $Level            = [LogLevel]::Info,
        [bool]           $IncludeTimestamp = $true,
        [bool]           $IncludeLevel     = $true,
        [string]         $Category         = ""
    )
    
    $script:LogConfig.Level = $Level
    $script:LogConfig.IncludeTimestamp = $IncludeTimestamp
    $script:LogConfig.IncludeLevel = $IncludeLevel
    $script:LogConfig.Category = $Category
}

# ログ出力方法の変更
function Switch-LogMethod {
    param(
        [LogOutputMethod]$Method
    )
    
    Initialize-Logger -Method $Method -Level $script:LogConfig.Level
}



# 初期化は各スクリプトで明示的に行う
# Initialize-Logger -Method ([LogOutputMethod]::Native) -Level ([LogLevel]::Info) 