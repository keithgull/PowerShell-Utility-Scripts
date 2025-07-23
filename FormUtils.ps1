[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
# 必要なアセンブリをロード
Add-Type -AssemblyName System.Windows.Forms

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
    $label          = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point($locationLeft, $locationTop)
    $label.Size     = New-Object System.Drawing.Size($width, $height)
	$form.Controls.Add($label)
    return $label
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

    # フォームオブジェクトを作成
    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Location = New-Object System.Drawing.Point($locationLeft, $locationTop)
    $textbox.Size = New-Object System.Drawing.Size($width, $height)
    $textBox.Multiline = $multilineMode
    $textBox.ScrollBars = $scrollBarType
	$form.Controls.Add($textbox)
    return $textbox
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
    $comboBox.Location = New-Object System.Drawing.Point($locationLeft, $locationTop)
    $comboBox.Size     = New-Object System.Drawing.Size($width, $height)

    # 表示アイテムを追加
    $comboBox.Items.AddRange($items)

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
    $form.Controls.Add($comboBox)
    return $comboBox
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
    return Create-ComboBoxWithValues -items $items -itemValues $items `
        -locationLeft $locationLeft -locationTop $locationTop `
        -width $width -height $height -form $form
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
    $label = Create-Label -caption $labelText -labelLeft $labelLeft -labelTop $labelTop -labelWidth $labelWidth -labelHeight $labelHeight -form $form
    # テキストボックスオブジェクトを作成
    $textBox = Create-TextBox -textBoxLeft $textBoxLeft -textBoxTop $textBoxTop -textBoxWidth $textBoxWidth -textBoxHeight $textBoxHeight -form $form
    # ラベルとテキストボックスをフォームに追加
    $form.Controls.Add($label)
    $form.Controls.Add($textBox)

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
    $button          = New-Object System.Windows.Forms.Button
    $button.Text     = $btnText
    $button.Location = New-Object System.Drawing.Point($btnXPosition, $btnYPosition)
    $form.Controls.Add($button)
    return $button
}


function Create-LabelComboBoxPair {
    param (
        [string]$labelText,
        [string[]]$items,
        [object[]]$itemValues = $null,
        [int]$labelLeft,
        [int]$labelTop,
        [int]$labelWidth,
        [int]$labelHeight,
        [int]$comboLeft,
        [int]$comboTop,
        [int]$comboWidth,
        [int]$comboHeight,
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

    return @($label, $comboBox)
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

