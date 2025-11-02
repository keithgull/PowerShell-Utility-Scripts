# PowerShell用ユーティリティモジュール
# 使用方法: . "c:\work\PSUtils.ps1"

# PSLogger.ps1の読み込み
# 注意: このモジュールはPSLogger.ps1に依存しています
# 使用方法: . "c:\work\PSUtils.ps1"  # PSLogger.ps1も自動的に読み込まれます
# 個別に読み込む場合: . "c:\work\PSLogger.ps1"; . "c:\work\PSUtils.ps1"
. "c:\work\PSLogger.ps1"

# エラーハンドリング付きで処理を実行する関数
# param1 実行するスクリプトブロック
# param2 エラー時のメッセージ（オプション）
# param3 エラー時のログレベル（デフォルト: Warn）
# param4 エラー時にnullを返すかどうか（デフォルト: true）
# 戻り値: スクリプトブロックの実行結果、エラー時はnullまたは例外を再スロー
function Invoke-WithErrorHandling {
    param(
        [scriptblock]$ScriptBlock,
        [string]$ErrorMessage = "",
        [LogLevel]$ErrorLevel = [LogLevel]::Warn,
        [bool]$ReturnNullOnError = $true
    )
    
    try {
        return & $ScriptBlock
    } catch {
        $actualMessage = if (Test-IsNullOrEmpty -Value $ErrorMessage) {
            "処理の実行中にエラーが発生しました: $($_.Exception.Message)"
        } else {
            $ErrorMessage
        }
        
        Write-PSLog -Level $ErrorLevel -Message $actualMessage
        
        if ($ReturnNullOnError) {
            return $null
        } else {
            throw
        }
    }
}

# フォルダ内の最新ファイルを取得する関数
# param1 対象フォルダパス
# param2 ファイル拡張子フィルタ
# param3 ログ出力を行うかどうか（デフォルト: true）
# 戻り値: 最新ファイルの名前（見つからない場合はnull）
function Get-LatestFile {
    param(
        [string]$folderPath,
        [string]$fileExtension,
        [bool]$enableLogging = $true
    )
    
    # ログ出力が有効な場合、パスと拡張子の情報を出力
    if ($enableLogging) {
        Write-PSLog -Level ([LogLevel]::Info) -Message "最新ファイル検索開始"
        Write-PSLog -Level ([LogLevel]::Info) -Message "  フォルダ: $folderPath"
        Write-PSLog -Level ([LogLevel]::Info) -Message "  フォルダ: 拡張子: $fileExtension"
    }
    
    # フォルダ内のファイルを更新日時順で取得
    $latestFile = Invoke-WithErrorHandling -ScriptBlock {
        Get-ChildItem -Path $folderPath -Filter $fileExtension | 
        Sort-Object LastWriteTime -Descending | 
        Select-Object -First 1
    } -ErrorMessage "フォルダ内のファイル取得に失敗しました" -ErrorLevel ([LogLevel]::Warn)
    
    # ファイルが見つかった場合の処理
    if ($null -ne $latestFile) {
        if ($enableLogging) {
            Write-PSLog -Level ([LogLevel]::Info) -Message "最新ファイルをデフォルトに設定: $($latestFile.Name) (更新日時: $($latestFile.LastWriteTime))"
            Write-PSLog -Level ([LogLevel]::Debug) -Message "Get-LatestFile: 最新ファイルを返します: $($latestFile.Name)"
        }
        return $latestFile.Name
    } else {
        if ($enableLogging) {
            Write-PSLog -Level ([LogLevel]::Warn) -Message "フォルダ内に$fileExtensionファイルが見つかりません"
            Write-PSLog -Level ([LogLevel]::Debug) -Message "Get-LatestFile: ファイルが見つかりません"
        }
        return $null
    }
}

# パス操作のユーティリティ関数群

# 安全なパス結合を行う関数
# param1 ベースパス
# param2 結合するパス
# 戻り値: 結合されたパス
function Join-PathSafely {
    param(
        [string]$BasePath,
        [string]$ChildPath
    )
    
    if (Test-IsNullOrEmpty -Value $BasePath) {
        return $ChildPath
    }
    if (Test-IsNullOrEmpty -Value $ChildPath) {
        return $BasePath
    }
    
    return Join-Path -Path $BasePath -ChildPath $ChildPath
}

# パスが存在するかチェックし、存在しない場合は代替パスを返す関数
# param1 チェック対象のパス
# param2 代替パス（配列で複数指定可能）
# 戻り値: 存在するパス、すべて存在しない場合は最初のパス
function Get-ValidPath {
    param(
        [string]$PrimaryPath,
        [string[]]$FallbackPaths = @()
    )
    
    # 最初にプライマリパスをチェック
    if (-not (Test-IsNullOrEmpty -Value $PrimaryPath) -and (Test-Path -Path $PrimaryPath)) {
        return $PrimaryPath
    }
    
    # フォールバックパスを順次チェック
    foreach ($path in $FallbackPaths) {
        if (-not (Test-IsNullOrEmpty -Value $path) -and (Test-Path -Path $path)) {
            return $path
        }
    }
    
    # すべて存在しない場合は最初のパスを返す
    return $PrimaryPath
}

# スクリプトのディレクトリパスを取得する関数
# 戻り値: スクリプトのディレクトリパス
function Get-ScriptDirectory {
    $scriptPath = $MyInvocation.MyCommand.Path
    if (Test-IsNullOrEmpty -Value $scriptPath) {
        $scriptPath = $PSScriptRoot
    }
    if (Test-IsNullOrEmpty -Value $scriptPath) {
        $scriptPath = Get-Location
    }
    
    return Split-Path -Parent $scriptPath
}

# パラメータ検証のユーティリティ関数群

# パラメータにデフォルト値を設定する関数
# param1 パラメータ値
# param2 デフォルト値
# 戻り値: パラメータ値（nullまたは空の場合はデフォルト値）
function Test-EmptyDefault {
    param(
        [string]$Value,
        [string]$DefaultValue
    )
    
    if (Test-IsNullOrEmpty -Value $Value) {
        return $DefaultValue
    }
    return $Value
}

# 数値パラメータの範囲を検証する関数
# param1 検証対象の値
# param2 最小値
# param3 最大値
# param4 デフォルト値
# 戻り値: 検証済みの値
function Test-NumberRange {
    param(
        [int]$Value,
        [int]$MinValue,
        [int]$MaxValue,
        [int]$DefaultValue
    )
    
    if ($Value -lt $MinValue -or $Value -gt $MaxValue) {
        Write-PSLog -Level ([LogLevel]::Warn) -Message "値 $Value が範囲外です。デフォルト値 $DefaultValue を使用します。"
        return $DefaultValue
    }
    return $Value
}

# 列挙値の検証を行う関数
# param1 検証対象の値
# param2 有効な値の配列
# param3 デフォルト値
# 戻り値: 検証済みの値
function Test-EnumValue {
    param(
        [string]$Value,
        [string[]]$ValidValues,
        [string]$DefaultValue
    )
    
    if (Test-IsNullOrEmpty -Value $Value) {
        return $DefaultValue
    }
    
    if ($Value -in $ValidValues) {
        return $Value
    } else {
        Write-PSLog -Level ([LogLevel]::Warn) -Message "無効な値 '$Value' が指定されました。デフォルト値 '$DefaultValue' を使用します。"
        return $DefaultValue
    }
}

# 対話型入力でデフォルト値付きの入力を取得する関数
#
# 引数:
#   PromptLabel: プロンプトのラベル名（{1}に当てはまる値、例: "リモート名"）
#   DefaultValue: デフォルト値（{2}に当てはまる値、必須）
#   PromptTemplate: プロンプトメッセージのテンプレート（省略時: "{1}を入力してください。（Enterでデフォルト：{2}を使用）"）
#                    {1}はPromptLabel、{2}はDefaultValueに置き換えられます
#   PromptMessage: 旧形式のパラメータ（後方互換性のため、PromptTemplateとPromptLabelの代わりに使用可能）
#
# 戻り値:
#   [PSCustomObject] 以下のプロパティを持つオブジェクト:
#     - Value: 入力された値またはデフォルト値
#     - IsDefault: デフォルト値が使用された場合は $true、入力値が使用された場合は $false
#
# 使用例:
#   # 新しい形式（推奨）
#   $result = Get-UserInputWithDefault -PromptLabel "リモート名" `
#                                       -DefaultValue "myproject-origin"
#   $remoteName = $result.Value
#
#   # テンプレートをカスタマイズする場合
#   $result = Get-UserInputWithDefault -PromptLabel "フォルダパス" `
#                                       -DefaultValue "C:\Work" `
#                                       -PromptTemplate "{1}を指定してください（デフォルト：{2}）"
function Get-UserInputWithDefault {
    param(
        [string]$PromptLabel = "",
        [Parameter(Mandatory=$true)]
        [string]$DefaultValue,
        [string]$PromptTemplate = "{1}を入力してください。（Enterでデフォルト：{2}を使用）",
        [string]$PromptMessage = ""
    )
    
    # 旧形式のPromptMessageが指定されている場合は、それを使用（後方互換性）
    $finalPromptMessage = ""
    if (-not [string]::IsNullOrWhiteSpace($PromptMessage)) {
        # 旧形式の場合、PromptMessageをそのまま使用
        $finalPromptMessage = $PromptMessage
        # PromptLabelが指定されていない場合は、PromptMessageから抽出
        if ([string]::IsNullOrWhiteSpace($PromptLabel)) {
            $PromptLabel = $PromptMessage -replace '[:：].*$', ''
        }
    } else {
        # 新しい形式：テンプレートに値を埋め込む
        if ([string]::IsNullOrWhiteSpace($PromptLabel)) {
            $PromptLabel = "値"
        }
        $finalPromptMessage = $PromptTemplate -replace '\{1\}', $PromptLabel -replace '\{2\}', $DefaultValue
    }
    
    # プロンプトメッセージを表示
    Write-Host ""
    Write-Host $finalPromptMessage -ForegroundColor Cyan
    $inputValue = Read-Host "$PromptLabel [$DefaultValue]"
    
    # 結果オブジェクトを作成
    $result = [PSCustomObject]@{
        Value     = ""
        IsDefault = $true
    }
    
    if ([string]::IsNullOrWhiteSpace($inputValue)) {
        # Enterキーが押された場合は、デフォルト値を使用
        $result.Value     = $DefaultValue
        $result.IsDefault = $true
        Write-Host "デフォルト値を使用します: $DefaultValue" -ForegroundColor Green
    } else {
        # 入力された場合は、その値を使用
        $result.Value     = $inputValue.Trim()
        $result.IsDefault = $false
        Write-Host "指定された値を使用します: $($result.Value)" -ForegroundColor Green
    }
    
    Write-Host ""
    return $result
} 