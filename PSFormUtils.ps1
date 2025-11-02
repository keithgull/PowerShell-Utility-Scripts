[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
# 必要なアセンブリをロード
Add-Type -AssemblyName System.Windows.Forms

# PSLogger.ps1の読み込み
# 注意: このモジュールはPSLogger.ps1に依存しています
# 使用方法: . "c:\work\PSLogger.ps1"; . "c:\work\PSFormUtils.ps1"
# または: . "c:\work\PSUtils.ps1"; . "c:\work\PSFormUtils.ps1"  # PSUtils経由でPSLoggerも読み込まれます
. "c:\work\PSLogger.ps1"

# コントロールのレイアウトを設定する関数
# param1 コントロールオブジェクト
# param2 左位置
# param3 上位置
# param4 幅
# param5 高さ
# 戻り値: 設定済みのコントロールオブジェクト
function Configure-ControlLayout {
    param(
        [System.Windows.Forms.Control]$control,
        [int]$left,
        [int]$top,
        [int]$width,
        [int]$height
    )
    
    $control.Location = New-Object System.Drawing.Point($left, $top)
    $control.Size     = New-Object System.Drawing.Size($width, $height)
    
    return $control
}

# フォームを作成する関数
# param1 フォームのタイトル
# param2 横幅
# param3 縦幅
#
function Create-Form {
    param (
        [string]$title,
        [int]$width,
        [int]$height
    )

    # フォームオブジェクトを作成
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $title
    $form.Size = New-Object System.Drawing.Size($width, $height)

    # フォームオブジェクトを返す
    return $form
}

function Create-Label {
    param (
        [string]$caption,
        [int]$locationLeft,
        [int]$locationTop,
        [int]$width,
        [int]$height,
        [System.Windows.Forms.Form]$form
    )

    # ラベルオブジェクトを作成
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $caption
    
    # レイアウトを設定
    $result = Configure-ControlLayout -control $label `
                            -left $locationLeft `
                            -top $locationTop `
                            -width $width `
                            -height $height
    
    # フォームに追加
    [void]$form.Controls.Add($label)
    
    # 戻り値が配列の場合は最初の要素（Label）を返す
    if ($result -is [array]) {
        $result = $result[0]
    }
    
    return $result
}

# テキストボックスを作成する。
# multilineMode
#  $true       : 複数行表示
#  $false      : １行表示
# scrollBarType
#  "None"      : スクロールバーなし
#  "Horizontal": 水平スクロールバー
#  "Vertical"  : 垂直スクロールバー
#  "Both"      : 両方表示
function Create-Text {
    param (
        [int]$locationLeft,
        [int]$locationTop,
        [int]$width,
        [int]$height,
        [bool]$multilineMode = $false,
        [string]$scrollBarType = "None",
        [System.Windows.Forms.Form]$form
     )

    # テキストボックスオブジェクトを作成
    $textbox = New-Object System.Windows.Forms.TextBox
    # レイアウトを設定
    $result = Configure-ControlLayout -control $textbox `
                                        -left $locationLeft `
                                        -top $locationTop `
                                        -width $width `
                                        -height $height
    $textbox.Multiline = $multilineMode
    
    # スクロールバーの設定
    switch ($scrollBarType) {
        "None"       { $textbox.ScrollBars = [System.Windows.Forms.ScrollBars]::None }
        "Horizontal" { $textbox.ScrollBars = [System.Windows.Forms.ScrollBars]::Horizontal }
        "Vertical"   { $textbox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical }
        "Both"       { $textbox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both }
        default      { $textbox.ScrollBars = [System.Windows.Forms.ScrollBars]::None }
    }
    
    # フォームに追加
    [void]$form.Controls.Add($textbox)
    
    # 戻り値が配列の場合は最初の要素（TextBox）を返す
    if ($result -is [array]) {
        $result = $result[0]
    }
    
    return $result
}

function Create-ComboBoxWithValues {
    param (
        [string[]]$items,            # 表示名の一覧
        [object[]]$itemValues,       # 対応する内部値（IDなど）
        [int]$locationLeft,
        [int]$locationTop,
        [int]$width,
        [int]$height,
        [System.Windows.Forms.Form]$form
    )

    if ($items.Count -ne $itemValues.Count) {
        throw "itemsとitemValuesの数が一致しません"
    }

    # ComboBox作成
    $comboBox = New-Object System.Windows.Forms.ComboBox
    # レイアウトを設定
    $result = Configure-ControlLayout -control $comboBox `
                            -left $locationLeft `
                            -top $locationTop `
                            -width $width `
                            -height $height

    # 表示アイテムを追加
    [void]$comboBox.Items.AddRange($items)

    # 初期選択
    if ($items.Length -gt 0) {
        $comboBox.SelectedIndex = 0
    }

    # 内部値対応表（表示名 → 値）を Tag に格納
    $mapping = @{}
    for ($i = 0; $i -lt $items.Count; $i++) {
        $mapping[$items[$i]] = $itemValues[$i]
    }
    $comboBox.Tag = $mapping

    # フォームに追加
    [void]$form.Controls.Add($comboBox)
    
    # 戻り値が配列の場合は最初の要素（ComboBox）を返す
    if ($result -is [array]) {
        $result = $result[0]
    }
    
    return $result
}

function Create-ComboBox {
    param (
        [string[]]$items,
        [int]$locationLeft,
        [int]$locationTop,
        [int]$width,
        [int]$height,
        [System.Windows.Forms.Form]$form
    )

        # 内部値として表示名そのまま渡す
    $result = Create-ComboBoxWithValues -items $items -itemValues $items `
        -locationLeft $locationLeft -locationTop $locationTop `
        -width $width -height $height -form $form
    
    # 戻り値が配列の場合は最初の要素（ComboBox）を返す
    if ($result -is [array]) {
        $result = $result[0]
    }
    
    return $result
}

# ラベルとテキストボックスのペアを作成する関数
function Create-LabelTextBoxPair {
    param (
        [string]$labelText,
        [int]$labelLeft,
        [int]$labelTop,
        [int]$labelWidth,
        [int]$labelHeight,
        [int]$textBoxLeft,
        [int]$textBoxTop,
        [int]$textBoxWidth,
        [int]$textBoxHeight,
        [System.Windows.Forms.Form]$form
    )

    # ラベルオブジェクトを作成
    $label = Create-Label -caption $labelText -locationLeft $labelLeft -locationTop $labelTop -width $labelWidth -height $labelHeight -form $form
    # テキストボックスオブジェクトを作成
    $textBox = Create-Text -locationLeft $textBoxLeft -locationTop $textBoxTop -width $textBoxWidth -height $textBoxHeight -form $form

    # ラベルとテキストボックスのオブジェクトを配列で返す
    return @($label, $textBox)
}

function Create-Button {
    param (
        [string]$btnText,
        [int]$btnXPosition,
        [int]$btnYPosition,
        [System.Windows.Forms.Form]$Form
    )
    # ボタンを作成
    $button = New-Object System.Windows.Forms.Button
    # レイアウトを設定（ボタンのデフォルトサイズを使用）
    $result = Configure-ControlLayout -control $button `
                            -left $btnXPosition `
                            -top $btnYPosition `
                            -width 75 `
                            -height 23
    $button.Text = $btnText
    
    # フォームに追加
    [void]$Form.Controls.Add($button)
    
    # 戻り値が配列の場合は最初の要素（Button）を返す
    if ($result -is [array]) {
        $result = $result[0]
    }
    
    return $result
}


# ラベルとコンボボックスのペアを作成する関数
# param1 ラベルのテキスト
# param2 コンボボックスに表示するアイテム配列
# param3 コンボボックスの内部値配列（オプション）
# param4 ラベルの左位置
# param5 ラベルの上位置
# param6 ラベルの幅
# param7 ラベルの高さ
# param8 コンボボックスの左位置
# param9 コンボボックスの上位置
# param10 コンボボックスの幅
# param11 コンボボックスの高さ
# param12 デフォルト選択アイテム（オプション、指定がない場合は配列の最初のアイテム）
# param13 フォームオブジェクト
# 戻り値: ラベルとコンボボックスの配列
function Create-LabelComboBoxPair {
    param (
        [string]         $labelText,
        [string[]]       $items,
        [object[]]       $itemValues = $null,
        [int]            $labelLeft,
        [int]            $labelTop,
        [int]            $labelWidth,
        [int]            $labelHeight,
        [int]            $comboLeft,
        [int]            $comboTop,
        [int]            $comboWidth,
        [int]            $comboHeight,
        [string]         $defaultItem = $null,
        [System.Windows.Forms.Form]$form
    )

    # ラベルを生成（既存のCreate-Labelを使用）
    $label = Create-Label `
        -caption 	  $labelText `
        -locationLeft $labelLeft `
        -locationTop  $labelTop `
        -width 		  $labelWidth `
        -height 	  $labelHeight `
        -form 		  $form

    # ComboBoxを生成
    if ($null -eq $itemValues) {
        $comboBox = Create-ComboBox `
            -items 		  $items `
            -locationLeft $comboLeft `
            -locationTop  $comboTop `
            -width 		  $comboWidth `
            -height 	  $comboHeight `
            -form 		  $form
    } else {
        $comboBox = Create-ComboBoxWithValues `
            -items 		  $items `
            -itemValues   $itemValues `
            -locationLeft $comboLeft `
            -locationTop  $comboTop `
            -width 		  $comboWidth `
            -height 	  $comboHeight `
            -form 		  $form
    }
    
    # ComboBoxオブジェクトの確認
    if ($null -eq $comboBox) {
        Write-PSLog -Level ([LogLevel]::Error) -Message "ComboBoxオブジェクトの作成に失敗しました"
        return @($label, $null)
    }
    
    Write-PSLog -Level ([LogLevel]::Debug) -Message "ComboBoxオブジェクト作成完了 - 型: $($comboBox.GetType().Name), アイテム数: $($comboBox.Items.Count)"

    # デフォルト選択アイテムを設定
    Write-PSLog -Level ([LogLevel]::Debug) -Message "Select-ComboItem呼び出し前 - ComboBox型: $($comboBox.GetType().Name), 値: $comboBox"
    $selectionResult = Select-ComboItem -comboBox $comboBox -defaultItem $defaultItem
    if (-not $selectionResult) {
        Write-PSLog -Level ([LogLevel]::Warn) -Message "ComboBoxのデフォルト選択設定に失敗しました"
    }

    return @($label, $comboBox)
}


# ComboBoxに指定されたアイテムが含まれているかチェックする関数
# param1 ComboBoxオブジェクト
# param2 チェック対象のアイテム
# 戻り値: アイテムが含まれている場合はtrue、含まれていない場合はfalse
function Test-ComboItemExists {
    param(
        [System.Windows.Forms.ComboBox]$comboBox,
        [string]$item
    )
    
    # ComboBoxオブジェクトの確認
    if ($null -eq $comboBox) {
        Write-PSLog -Level ([LogLevel]::Error) -Message "ComboBoxオブジェクトがnullです"
        return $false
    }
    
    # アイテムの確認
    if (Test-IsNullOrEmpty -Value $item) {
        Write-PSLog -Level ([LogLevel]::Warn) -Message "チェック対象のアイテムがnullまたは空です"
        return $false
    }
    
    # ComboBoxからアイテムリストを取得
    $items = @($comboBox.Items)
    
    # アイテムが含まれているかチェック
    $exists = $item -in $items
    
    if ($exists) {
        Write-PSLog -Level ([LogLevel]::Debug) -Message "アイテムが存在します: $item"
    } else {
        Write-PSLog -Level ([LogLevel]::Debug) -Message "アイテムが存在しません: $item"
    }
    
    return $exists
}

# ComboBoxのデフォルト選択アイテムを設定する関数
# param1 ComboBoxオブジェクト
# param2 デフォルト選択アイテム（オプション、指定がない場合は最初のアイテムを選択）
# 戻り値: 設定成功時はtrue、失敗時はfalse
#
# 処理フロー:
# Select-ComboItem
# ├── defaultItemが指定されている場合
# │   ├── Test-ComboItemExistsで存在チェック
# │   └── _Set-ComboItemByValueで設定
# │       ├── SelectedItemで設定を試行
# │       └── 失敗時は_Set-ComboItemByIndexでフォールバック
# └── defaultItemが指定されていない場合
#     └── _Set-ComboItemByIndexで最初のアイテムを設定
function Select-ComboItem {
    param(
        [System.Windows.Forms.ComboBox]$comboBox,
        [string]$defaultItem = $null
    )
    
    # ComboBoxオブジェクトの確認
    if ($null -eq $comboBox) {
        Write-PSLog -Level ([LogLevel]::Error) -Message "ComboBoxオブジェクトがnullです"
        return $false
    }
    
    # ComboBoxからアイテムリストを取得
    $items = @($comboBox.Items)
    
    # コントロールの初期化を待つ
    Start-Sleep -Milliseconds 50
    
    if ($null -ne $defaultItem -and $defaultItem -ne "") {
        # 指定されたデフォルトアイテムがアイテムリストに含まれているかチェック
        if (Test-ComboItemExists -comboBox $comboBox -item $defaultItem) {
            return _Set-ComboItemByValue -comboBox $comboBox -items $items -defaultItem $defaultItem
        } else {
            Write-PSLog -Level ([LogLevel]::Warn) -Message "デフォルトアイテムが見つかりません: $defaultItem"
            return $false
        }
    } else {
        # デフォルトアイテムが指定されていない場合は配列の最初のアイテムを選択
        return _Set-ComboItemByIndex -comboBox $comboBox -items $items -index 0
    }
}

# ComboBoxに指定されたアイテムをSelectedItemで設定する内部関数
# param1 ComboBoxオブジェクト
# param2 アイテムリスト
# param3 設定するアイテム
# 戻り値: 設定成功時はtrue、失敗時はfalse
function _Set-ComboItemByValue {
    param(
        [System.Windows.Forms.ComboBox]$comboBox,
        [string[]]$items,
        [string]$defaultItem
    )
    
    try {
        $comboBox.SelectedItem = $defaultItem
        Write-PSLog -Level ([LogLevel]::Debug) -Message "デフォルトアイテムを設定しました: $defaultItem"
        return $true
    } catch {
        Write-PSLog -Level ([LogLevel]::Warn) -Message "ComboBoxのSelectedItem設定に失敗しました: $($_.Exception.Message)"
        # フォールバック: SelectedIndexを使用
        $index = [array]::IndexOf($items, $defaultItem)
        if ($index -ge 0) {
            return _Set-ComboItemByIndex -comboBox $comboBox -items $items -index $index
        }
        return $false
    }
}

# ComboBoxに指定されたインデックスのアイテムを設定する内部関数
# param1 ComboBoxオブジェクト
# param2 アイテムリスト
# param3 設定するインデックス
# 戻り値: 設定成功時はtrue、失敗時はfalse
function _Set-ComboItemByIndex {
    param(
        [System.Windows.Forms.ComboBox]$comboBox,
        [string[]]$items,
        [int]$index
    )
    
    # インデックスの範囲チェック
    if ($index -lt 0 -or $index -ge $items.Count) {
        Write-PSLog -Level ([LogLevel]::Warn) -Message "インデックスが範囲外です: $index (アイテム数: $($items.Count))"
        return $false
    }
    
    try {
        Start-Sleep -Milliseconds 25
        $comboBox.SelectedIndex = $index
        Write-PSLog -Level ([LogLevel]::Debug) -Message "インデックスでアイテムを設定しました: $($items[$index]) (インデックス: $index)"
        return $true
    } catch {
        Write-PSLog -Level ([LogLevel]::Warn) -Message "ComboBoxのSelectedIndex設定に失敗しました: $($_.Exception.Message)"
        return $false
    }
}

# ラベルとラジオボタングループを作成する関数
# param1 ラベルテキスト（オプション、グループ全体の説明ラベルとして使用）
# param2 ラジオボタンのラベル名配列
# param3 ラジオボタングループの左位置（グループ全体の説明ラベルがない場合はラジオボタンの開始位置、ある場合はラベルの左位置）
# param4 ラジオボタングループの上位置（グループ全体の説明ラベルがない場合はラジオボタンの開始位置、ある場合はラベルの上位置）
# param5 ラジオボタンの◯をラベルテキストの先頭に置くか後ろに置くかのフラグ（"Before": 先頭（◯ ラベル）, "After": 後ろ（ラベル ◯）、デフォルト: "Before"）
# param6 レイアウトを縦にするか横にするかのフラグ（"Vertical": 縦, "Horizontal": 横、デフォルト: "Vertical"）
# param7 ラベルテキストがある場合のラベルの左位置（省略可、指定がない場合はgroupLeftを使用）
# param8 ラベルテキストがある場合のラベルの上位置（省略可、指定がない場合はgroupTopを使用）
# param9 ラベルテキストがある場合のラベルの幅（省略可、デフォルト: 150）
# param10 ラベルテキストがある場合のラベルの高さ（省略可、デフォルト: 20）
# param11 フォームオブジェクト
# 戻り値: ハッシュテーブル @{ Label = グループ全体の説明ラベルオブジェクト（存在する場合）, RadioButtons = ラジオボタン配列, Panel = FlowLayoutPanelオブジェクト }
#
# 処理フロー:
# Create-LabelRadioButtonGroup
# ├── FlowLayoutPanelを作成してレイアウト方向を設定
# ├── グループ全体の説明ラベルが指定されている場合
# │   └── ラベルを作成し、ラジオボタンパネルの位置を調整
# ├── ラジオボタンを順次作成
# │   ├── radioButtonPositionが"Before"の場合: 標準の位置（◯が左、テキストが右）
# │   └── radioButtonPositionが"After"の場合: RightToLeftを使用して◯を右、テキストを左に配置
# └── 最初のラジオボタンを選択状態に設定
function Create-LabelRadioButtonGroup {
    param (
        [string]                                $labelText                   = $null,
        [string[]]                              $radioButtonLabels,
        [int]                                   $groupLeft,
        [int]                                   $groupTop,
        [ValidateSet("Before", "After")]        [string]$radioButtonPosition = "Before",
        [ValidateSet("Vertical", "Horizontal")] [string]$layoutDirection     = "Vertical",
        [int]                                   $labelLeft                   = $null,
        [int]                                   $labelTop                    = $null,
        [int]                                   $labelWidth                  = 150,
        [int]                                   $labelHeight                 = 20,
        [System.Windows.Forms.Form]             $form
    )
    
    # ラジオボタンラベル配列の確認
    if ($null -eq $radioButtonLabels -or $radioButtonLabels.Count -eq 0) {
        Write-PSLog -Level ([LogLevel]::Error) -Message "ラジオボタンのラベル配列が空です"
        return $null
    }
    
    # グループ全体の説明ラベルオブジェクト（存在する場合）
    $label = $null
    
    # パネルの初期位置を決定
    $panelLeft = $groupLeft
    $panelTop  = $groupTop
    
    # グループ全体の説明ラベルが指定されている場合、ラベルを作成してパネル位置を調整
    if ($null -ne $labelText -and $labelText -ne "") {
        # ラベルの位置が指定されていない場合は、グループ位置を使用
        $actualLabelLeft = $labelLeft
        $actualLabelTop = $labelTop
        if ($null -eq $actualLabelLeft) {
            $actualLabelLeft = $groupLeft
        }
        if ($null -eq $actualLabelTop) {
            $actualLabelTop = $groupTop
        }
        
        # ラベルを作成
        $label = Create-Label `
            -caption      $labelText `
            -locationLeft $actualLabelLeft `
            -locationTop  $actualLabelTop `
            -width        $labelWidth `
            -height       $labelHeight `
            -form         $form
        
        # パネルの位置をラベルの横に配置（縦レイアウトでも横レイアウトでも同じ行から開始）
        # ラベルの右側、同じY位置に配置
        $panelLeft = $actualLabelLeft + $labelWidth + 10
        $panelTop  = $actualLabelTop
    }
    
    # FlowLayoutPanelを作成
    $panel = New-Object System.Windows.Forms.FlowLayoutPanel
    $panel.Location = New-Object System.Drawing.Point($panelLeft, $panelTop)
    # 折り返しを機能させるため、幅を制限する（AutoSizeはfalseにする）
    # フォームの幅からパネルの左位置とマージンを引いて計算
    $maxWidth = $form.Width - $panelLeft - 40
    $panel.Width = $maxWidth
    $panel.AutoSize = $false
    $panel.Height = 200  # 初期高さ（自動調整されないため、後で調整が必要な場合は大きく設定）
    # 自動折り返しを有効にする（項目数が多い場合に複数行/複数列に折り返す）
    $panel.WrapContents = $true
    
    # レイアウト方向を設定
    if ($layoutDirection -eq "Vertical") {
        # 縦レイアウト：上から下に配置し、右に折り返す（複数列になる）
        $panel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
        # 縦レイアウトの場合、ラジオボタン間の間隔を設定
        $panel.Padding = New-Object System.Windows.Forms.Padding(0, 0, 15, 5)
    } else {
        # 横レイアウト：左から右に配置し、下に折り返す（複数行になる）
        $panel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
        # 横レイアウトの場合、ラジオボタン間の間隔を設定
        $panel.Padding = New-Object System.Windows.Forms.Padding(0, 0, 10, 5)
    }
    
    # ラジオボタンを作成してパネルに追加
    $radioButtons = @()
    foreach ($radioLabel in $radioButtonLabels) {
        $radioButton = New-Object System.Windows.Forms.RadioButton
        $radioButton.Text = $radioLabel
        $radioButton.AutoSize = $true
        
        # ラジオボタンの◯をラベルテキストの後ろに配置する場合（ラベル ◯ の形式）
        if ($radioButtonPosition -eq "After") {
            # RightToLeftを使用して◯を右側、テキストを左側に配置
            $radioButton.RightToLeft = [System.Windows.Forms.RightToLeft]::Yes
        }
        # radioButtonPositionが"Before"の場合は標準の位置（◯ ラベル）で配置される
        
        $radioButtons += $radioButton
        [void]$panel.Controls.Add($radioButton)
    }
    
    # 最初のラジオボタンを選択状態にする
    if ($radioButtons.Count -gt 0) {
        $radioButtons[0].Checked = $true
    }
    
    # パネルをフォームに追加
    [void]$form.Controls.Add($panel)
    
    # 戻り値をハッシュテーブルで返す
    return @{
        Label        = $label
        RadioButtons = $radioButtons
        Panel        = $panel
    }
}


# Enumのようなクラスを定義
class ButtonType {
    static [string]$psOK = "psOK"
    static [string]$psOKCancel = "psOKCancel"
    static [string]$psYesNo = "psYesNo"
    static [string]$psYesNoCancel = "psYesNoCancel"
}

function Create-FormButtons {
    param (
        [string]$ButtonType,
        [int]$XPosition,
        [int]$YPosition,
        [System.Windows.Forms.Form]$Form
    )

    $buttonWidth  = 50
    $buttonHeight = 30
    $spacing = 10
    $buttons = @()

    switch ($ButtonType) {
        [ButtonType]::psOK {
            $buttons += New-Object System.Windows.Forms.Button
            $buttons.Text = "OK"
        }
        [ButtonType]::psOKCancel {
            $buttons += New-Object System.Windows.Forms.Button
            $buttons.Text = "OK"
            $buttons += New-Object System.Windows.Forms.Button
            $buttons.Text = "Cancel"
        }
        [ButtonType]::psYesNo {
            $buttons += New-Object System.Windows.Forms.Button
            $buttons.Text = "Yes"
            $buttons += New-Object System.Windows.Forms.Button
            $buttons.Text = "No"
        }
        [ButtonType]::psYesNoCancel {
            $buttons += New-Object System.Windows.Forms.Button
            $buttons.Text = "Yes"
            $buttons += New-Object System.Windows.Forms.Button
            $buttons.Text = "No"
            $buttons += New-Object System.Windows.Forms.Button
            $buttons.Text = "Cancel"
        }
    }

    for ($i = 0; $i -lt $buttons.Count; $i++) {
        $buttons[$i].Width = $buttonWidth
        $buttons[$i].Height = $buttonHeight
        $buttons[$i].Left = $XPosition + ($i * ($buttonWidth + $spacing))
        $buttons[$i].Top = $YPosition
        $Form.Controls.Add($buttons[$i])
    }

    return $buttons
}


# メッセージダイアログを表示する関数
# 使用例
# $resultMessage = Show-MessageDialog -Message "これはテストメッセージです。" -Title "テストタイトル" -ButtonType "YesNo"
#
function Show-MessageDialog {
    param (
        [string]$Message,
        [string]$Title,
        [ValidateSet("OK", "OKCancel", "YesNo", "YesNoCancel", "RetryCancel", "AbortRetryIgnore")]
        [string]$ButtonType
    )

    # ボタンのタイプを対応するMessageBoxButtonsに変換
    switch ($ButtonType) {
        "OK"               { $buttons = [System.Windows.Forms.MessageBoxButtons]::OK }
        "OKCancel"         { $buttons = [System.Windows.Forms.MessageBoxButtons]::OKCancel }
        "YesNo"            { $buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo }
        "YesNoCancel"      { $buttons = [System.Windows.Forms.MessageBoxButtons]::YesNoCancel }
        "RetryCancel"      { $buttons = [System.Windows.Forms.MessageBoxButtons]::RetryCancel }
        "AbortRetryIgnore" { $buttons = [System.Windows.Forms.MessageBoxButtons]::AbortRetryIgnore }
    }

    # メッセージダイアログを表示し、押されたボタンを取得
    $result = [System.Windows.Forms.MessageBox]::Show($Message, $Title, $buttons)

	return $result
}

# 結果表示用ダイアログフォームを作成する関数
# param1 フォームのタイトル
# param2 表示するテキスト内容
# param3 テキストエリアの編集可否（デフォルト: false）
# param4 フォームの横幅（デフォルト: 600）
# param5 フォームの縦幅（デフォルト: 400）
function Show-ResultDialog {
    param (
        [string]$title   = "",
        [string]$content = "",
        [bool]$readOnly  = $false,
        [int]$width      = 600,
        [int]$height     = 400
    )

    # フォームの作成
    $form = Create-Form -title $title -width $width -height $height

    # テキストエリアの作成（複数行、垂直スクロールバー付き）
    $textArea = Create-Text `
        -locationLeft 20 `
        -locationTop 20 `
        -width ($width - 40) `
        -height ($height - 100) `
        -multilineMode $true `
        -scrollBarType "Vertical" `
        -form $form

    # テキストエリアの編集可否を設定
    $textArea.ReadOnly = $readOnly
    $textArea.Text = $content

    # ボタンの作成
    $copyButton = Create-Button -btnText "コピー" -btnXPosition 20 -btnYPosition ($height - 60) -Form $form
    $closeButton = Create-Button -btnText "閉じる" -btnXPosition 120 -btnYPosition ($height - 60) -Form $form

    # コピーボタンのクリックイベント処理
    $copyButton.Add_Click({
        try {
            [System.Windows.Forms.Clipboard]::SetText($textArea.Text)
            [System.Windows.Forms.MessageBox]::Show("クリップボードにコピーしました。", "コピー完了", "OK", "Information")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("クリップボードへのコピーに失敗しました。", "エラー", "OK", "Error")
        }
    })

    # 閉じるボタンのクリックイベント処理
    $closeButton.Add_Click({
        $form.Close()
    })

    # フォームを表示
    $form.Topmost = $true
    [void]$form.ShowDialog()

    # フォームを解放
    $form.Dispose()
}

