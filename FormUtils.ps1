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
        [int]$locationTop
        [int]$width,
        [int]$height,
        [System.Windows.Forms.Form]$form
    )

    # ラベルオブジェクトを作成
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point($locationLeft, $locationTop)
    $label.Size = New-Object System.Drawing.Size($width, $height)
	$form.Controls.Add($label)
    return $label
}

function Create-Text {
    param (
        [int]$locationLeft,
        [int]$locationTop
        [int]$width,
        [int]$height,
        [System.Windows.Forms.Form]$form
     )

    # フォームオブジェクトを作成
    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Location = New-Object System.Drawing.Point($locationLeft, $locationTop)
    $textbox.Size = New-Object System.Drawing.Size($width, $height)
	$form.Controls.Add($textbox)
    return $textbox
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

