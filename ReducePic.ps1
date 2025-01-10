#-------------------------------------------------------------------------------
#  ReducePic.ps1
#  フォルダー内の画像ファイルを一括で縮小する
#    幅と高さの長い方が指定したサイズより大きい画像を、長い方が指定したサイズになるように縦横比を維持して縮小し、
#    出力先フォルダーに出力する。ファイル名、Exif情報、ファイル作成日時/更新日時はそのまま保持する。
#    幅と高さのいずれも指定したサイズ以下の画像は何もせずそのまま出力先にコピーする。
#-------------------------------------------------------------------------------

# ファイル/フォルダー選択ダイアログ、メッセージボックス等を使用するためのアセンブリのロード
Add-Type -AssemblyName System.Windows.Forms

# GDI+ の基本的なグラフィックス機能
Add-Type -AssemblyName System.Drawing

# WPF(Windows Presentation Foundation)の機能を使用するためのアセンブリのロード
# WPFのメッセージボックスを使用するため（PowerShell ISE環境ではデフォルトでロードされているが、そうでなければ必要）
Add-Type -AssemblyName PresentationFramework

$ErrorActionPreference = 'Stop' # このスクリプト内でのエラー発生時に処理を中断する（PowerShellの既定は Continue:エラーメッセージを出して処理継続）

# 画像のリサイズ（ここでは縮小限定）
function resizeImage {
  param (
    [string] $inputFile, # 入力ファイル（フルパス）
    [string] $outputFile, # 出力ファイル（フルパス）
    [string] $picExt1, # 画像形式
    [int] $sizeLongSide, # リサイズ後の長辺の大きさ
    [int] $sizeRate # リサイズ割合(1%～99%)　sizeLongSideとsizeRateはどちらか一方のみ有効
  )

    # 画像ファイルの読み込み
    $image = [System.Drawing.Image]::FromFile($inputFile)

    #$msg = $inputFile + "`r`n" + $outputFile + "`r`n" + $image.Width + "`r`n" + $image.Height
    #$msgBoxResult = $msgBox::Show($dummyWindow, $msg, 'ReducePic', $buttonsOK, $iconInfo, $defaultButton)
    #$msgBoxResult = $msgBox::Show($dummyWindow, $saveFormat, 'ReducePic', $buttonsOK, $iconInfo, $defaultButton)

    [int] $newWidth = 0
    [int] $newHeight = 0
    if ($sizeLongSide -gt 0) {
        if ($image.Width -ge $image.Height -and $image.Width -gt $sizeLongSide) {
            $newWidth = $sizeLongSide
            $newHeight = [Math]::Round(($image.Height * $sizeLongSide / $image.Width), 0, [MidpointRounding]::AwayFromZero)
        } elseif ($image.Height -gt $image.Width -and $image.Height -gt $sizeLongSide) {
            $newHeight = $sizeLongSide
            $newWidth = [Math]::Round(($image.Width * $sizeLongSide / $image.Height), 0, [MidpointRounding]::AwayFromZero)
        } else {
            copy $inputFile $outputFile
            $image.Dispose()
            return
        }
    } elseif ($sizeRate -ge 1 -and $sizeRate -le 99) {
        $newWidth = [Math]::Round(($image.Width * $sizeRate / 100), 0, [MidpointRounding]::AwayFromZero)
        $newHeight = [Math]::Round(($image.Height * $sizeRate / 100), 0, [MidpointRounding]::AwayFromZero)
    } else {
        return
    }

    # 縮小先のオブジェクトを生成
    $canvas = New-Object System.Drawing.Bitmap ($newWidth, $newHeight)

    # 属性の引き継ぎ(Exifや回転情報など)
    foreach($prop in $image.PropertyItems) {
        $canvas.SetPropertyItem($prop);
    }

    # 縮小先のビットマップでGraphicオブジェクトを作成し、元図形を縮小先のビットマップの大きさで描画する
    $graphics = [System.Drawing.Graphics]::FromImage($canvas)
    $graphics.InterpolationMode = $interpolation  # リサイズ処理での補完方法を設定
    $graphics.DrawImage($image, 0, 0, $newWidth, $newHeight)

    switch ($picExt1) {
        '*.jpg' {
            $saveFormat = [System.Drawing.Imaging.ImageFormat]::Jpeg
            $mime = 'image/jpeg'
        }
        '*.png' {
            $saveFormat = [System.Drawing.Imaging.ImageFormat]::Png
            $mime = 'image/png'
        }
        '*.bmp' {
            $saveFormat = [System.Drawing.Imaging.ImageFormat]::Bmp
            $mime = 'image/bmp'
        }
        '*.gif' {
            $saveFormat = [System.Drawing.Imaging.ImageFormat]::Gif
            $mime = 'image/gif'
        }
        '*.tif' {
            $saveFormat = [System.Drawing.Imaging.ImageFormat]::Tiff
            $mime = 'image/tiff'
        }
    }

    # 保存
    #    JPEGの保存品質を指定する場合はSaveメソッドのEncoderParametersを用いるオーバーロードを使う必要がある。
    #    逆に品質を指定しない場合は、ImageFormatだけ用いるオーバーロードを使って品質は既定の動作に任せることにする。
    if ($picExt1 -eq "*.jpg" -and $jpegQuality -ge 0 -and $jpegQuality -le 100) {
        $encoderCategory = [System.Drawing.Imaging.Encoder]::Quality
        $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
        #$encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($encoderCategory, 75)
        $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($encoderCategory, $jpegQuality)
        $encoderInfo = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | where {$_.MimeType -eq $mime}
        $canvas.Save($outputFile, $encoderInfo, $encoderParams)

        $encoderParams.Param[0].Dispose()
        $encoderParams.Dispose()
    } else {
        #$canvas.Save($outputFile, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        $canvas.Save($outputFile, $saveFormat)
    }

    # オブジェクトの破棄
    $graphics.Dispose()
    $canvas.Dispose()
    $image.Dispose()

    # タイムスタンプの引き継ぎ
    $lastWriteTime=(Get-ItemProperty $inputFile).LastWriteTime
    $creationTime=(Get-ItemProperty $inputFile).CreationTime
    Set-ItemProperty $outputFile -name LastWriteTime -Value $lastWriteTime
    Set-ItemProperty $outputFile -name CreationTime -Value $creationTime
}




#$script:inputFolder = ""  # 子スコープから読み書きするのでscriptスコープで定義
#$script:outputFolder = ""  # 子スコープから読み書きするのでscriptスコープで定義

$msgBox = [System.Windows.MessageBox]
$buttonsOK=[System.Windows.MessageBoxButton]::OK
$defaultButton=[System.Windows.MessageBoxResult]::OK  # 既定の結果
$iconErr=[System.Windows.MessageBoxImage]::Error
$iconInfo=[System.Windows.MessageBoxImage]::Information
$dummyWindow = New-Object System.Windows.Window
$dummyWindow.Width = 100
$dummyWindow.Height = 100
$dummyWindow.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
$dummyWindow.Topmost = $true

#何らかの理由で設定ファイルから読み込めなかった場合のデフォルト値
$interpolation = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
[long] $jpegQuality = 999  # EncoderParameterで指定する品質の数値はlong(int64)型でなければエラーになる。
#設定ファイル読み込み
$CONFIG_FILE = './ReducePic.ini'
$PARAM = @{}  # 連想配列（ハッシュテーブル）の宣言
Get-Content $CONFIG_FILE | %{$PARAM += ConvertFrom-StringData $_}
if ($PARAM.ContainsKey('Interpolation')) {  # キーの存在チェック
    if ([int]::TryParse($PARAM.Interpolation, [ref]$null)) {  # 型チェック(int)
        switch ($PARAM.Interpolation) {
            '1' {  # 高品質バイキュービック法
                $interpolation = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
                $interpolationName = ''
            }
            '2' {  # 高品質バイリニア法
                $interpolation = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBilinear
                $interpolationName = '2:高品質バイリニア法'
            }
            '3' {  # バイキュービック法
                $interpolation = [System.Drawing.Drawing2D.InterpolationMode]::Bicubic
                $interpolationName = '3:バイキュービック法'
            }
            '4' {  # バイリニア法
                $interpolation = [System.Drawing.Drawing2D.InterpolationMode]::Bilinear
                $interpolationName = '4:バイリニア法'
            }
            '5' {  # ニアレストネイバー法
                $interpolation = [System.Drawing.Drawing2D.InterpolationMode]::NearestNeighbor
                $interpolationName = '5:ニアレストネイバー法'
            }
        }
    }
}
if ($PARAM.ContainsKey('Quality')) {  # キーの存在チェック
    if ($PARAM.Quality -eq 'X') {
        $jpegQuality = 999
    } elseif ([int]::TryParse($PARAM.Quality, [ref]$null)) {  # 型チェック(int)
        $jpegQuality = [long]$PARAM.Quality
        if ($jpegQuality -lt 0 -or $jpegQuality -gt 100) {
                $jpegQuality = 999
        }
    }
}
#[System.Windows.Forms.MessageBox]::Show($interpolation.value__, "Info")
#[System.Windows.Forms.MessageBox]::Show($jpegQuality, "Info")

#---- 高DPI(スケーリング)に対応したフォームの作成 ----
# [High DPI] 最初にSetProcessDPIAwareを実行して高DPIの対応を宣言する
Add-Type -TypeDefinition @'
using System.Runtime.InteropServices;

public class ProcessDPI {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern bool SetProcessDPIAware();      
}
'@ -ReferencedAssemblies 'System.Drawing.dll'
$null = [ProcessDPI]::SetProcessDPIAware()

# ビジュアルスタイルを有効にする
[System.Windows.Forms.Application]::EnableVisualStyles()

# フォーム作成
$form = New-Object System.Windows.Forms.Form

# [High DPI] AutoScaleModeのDpi指定を機能させるには（DPIに従ったスケーリングでコントロールが配置およびサイズ設定されるには）、
# フォームの作成直後にSuspendLayoutを呼び出し、すべてのコントロールを作成して配置したらResumeLayoutを呼び出す。
$form.SuspendLayout()
$form.AutoScaleDimensions =  New-Object System.Drawing.SizeF(96, 96)  # コントロールがデザインされたときのサイズを設定する
$form.AutoScaleMode  = [System.Windows.Forms.AutoScaleMode]::Dpi # 自動スケーリングのモード = Dpi:ディスプレイの解像度に応じてスケールを制御

#$form.Text = 'Resize all pictures in folder'
$form.Text = 'ReducePic  ─フォルダー内の画像ファイルを一括で縮小する─'
#$form.Size = New-Object System.Drawing.Size(300,200)  # フォーム全体のサイズで指定
$form.ClientSize = New-Object System.Drawing.Size(500,400)  # クライアント領域のサイズで指定（枠やタイトルバーを除いたサイズ）
$form.StartPosition = 'CenterScreen'
#$fontDefault = New-Object System.Drawing.Font("Yu Gothic UI",$form.Font.Size)
#$form.Font = $fontDefault

$label01 = New-Object System.Windows.Forms.Label
$label01.Location = New-Object System.Drawing.Point(150,20)
$label01.Size = New-Object System.Drawing.Size(340,20)
#$label01.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$label01.Text = '（選択するフォルダーの中に入った状態で[開く]を押して選択します）'
$label01.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$fontDefault = New-Object System.Drawing.Font("Yu Gothic UI",$label01.Font.Size)
$label01.Font = $fontDefault
$form.Controls.Add($label01)

$label_folder1 = New-Object System.Windows.Forms.Label
$label_folder1.Location = New-Object System.Drawing.Point(10,30)
$label_folder1.Size = New-Object System.Drawing.Size(200,22)
#$label_folder1.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$label_folder1.Text = '元画像のフォルダー'
$font_L = New-Object System.Drawing.Font("Yu Gothic UI",11)
$label_folder1.Font = $font_L
$form.Controls.Add($label_folder1)

# DPI Scalingの取得
$DPISetting = (Get-ItemProperty 'HKCU:\Control Panel\Desktop\WindowMetrics' -Name AppliedDPI).AppliedDPI
[float]$displayScale = $DPISetting / 96;  # 1.0,1.25,1.5…等の値にする

$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = New-Object System.Drawing.Point(10,52)
$textBox1.Size = New-Object System.Drawing.Size(405,60)
$textBox1.Multiline = $true
# 注：ISEから実行した場合は、テキストボックスだけなぜかAutoScaleが効かない。スケーリング係数をかけたフォントサイズを直指定すればサイズは合うが、それだと通常実行時に合わなくなる。
#$fontTextbox = New-Object System.Drawing.Font($textBox.Font.Name,(12 * $displayScale))
$fontTextbox = New-Object System.Drawing.Font($textBox1.Font.Name,12)
$textBox1.Font =  $fontTextbox
$form.Controls.Add($textBox1)

$browseButton1 = New-Object System.Windows.Forms.Button
$browseButton1.Location = New-Object System.Drawing.Point(420,52)
$browseButton1.Size = New-Object System.Drawing.Size(70,23)
$browseButton1.Text = '参照...'
$browseButton1.Font = $fontDefault
$form.Controls.Add($browseButton1)

# browseボタンのクリックイベント
$Browse1 = {
    <#
    #フォルダー選択ダイアログのインスタンス作成
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog

    $setPath = [Environment]::GetFolderPath("Personal") # "Personal"または"MyDocuments" : ユーザーのドキュメントフォルダー
    # RootFolderは選択できる上限のパス位置　SelectedPathは最初にデフォルトで選択されているパス（返ってきた後は実際に選択されたパス）
    #$dialog.RootFolder = "MyComputer"  # "Desktop","Personal"...
    $dialog.SelectedPath = $setPath
    $dialog.Description = "元画像があるフォルダーを指定してください"
    #>

    # ファイル選択用のOpenFileDialogを応用してフォルダーを選択する方法
    #ファイル選択ダイアログのインスタンス作成
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $setPath = [Environment]::GetFolderPath("Personal")
    if ($textBox1.Text) {  # nullでも空でもない場合
        $dialog.InitialDirectory = $textBox1.Text
    } else {
        $dialog.InitialDirectory = $setPath
    }
    $dialog.InitialDirectory = $setPath
    $dialog.Title = 'フォルダーを選択してください'
    $dialog.ValidateNames = 0
    $dialog.CheckFileExists = 0
    $dialog.CheckPathExists = 1
    $dialog.FileName = 'フォルダーを選択'
    
    #フォルダー選択ダイアログ表示
    $retDialog = $dialog.ShowDialog()
    if ($retDialog -ne [System.Windows.Forms.DialogResult]::OK) {
        # オブジェクトの後始末
        $dialog.Dispose()
        return
    }
    #[System.Windows.Forms.MessageBox]::Show($dialog.FileName)
    $textBox1.Text = Split-Path -Parent $dialog.FileName
}
$browseButton1.Add_Click($Browse1)

$label_folder2 = New-Object System.Windows.Forms.Label
$label_folder2.Location = New-Object System.Drawing.Point(10,120)
$label_folder2.Size = New-Object System.Drawing.Size(200,22)
$label_folder2.Text = '出力先画像のフォルダー'
$label_folder2.Font = $font_L
$form.Controls.Add($label_folder2)

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(10,142)
$textBox2.Size = New-Object System.Drawing.Size(405,60)
$textBox2.Multiline = $true
$textBox2.Font =  $fontTextbox
$form.Controls.Add($textBox2)

$browseButton2 = New-Object System.Windows.Forms.Button
$browseButton2.Location = New-Object System.Drawing.Point(420,142)
$browseButton2.Size = New-Object System.Drawing.Size(70,23)
$browseButton2.Text = '参照...'
$browseButton2.Font = $fontDefault
$form.Controls.Add($browseButton2)

$Browse2 = {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $setPath = [Environment]::GetFolderPath("Personal")
    if ($textBox2.Text) {  # nullでも空でもない場合
        $dialog.InitialDirectory = $textBox2.Text
    } elseif ($textBox1.Text) {
        $dialog.InitialDirectory = $textBox1.Text
    } else {
        $dialog.InitialDirectory = $setPath
    }
    $dialog.Title = 'フォルダーを選択してください'
    $dialog.ValidateNames = 0
    $dialog.CheckFileExists = 0
    $dialog.CheckPathExists = 1
    $dialog.FileName = 'フォルダーを選択'
    
    #フォルダー選択ダイアログ表示
    $retDialog = $dialog.ShowDialog()
    if ($retDialog -ne [System.Windows.Forms.DialogResult]::OK) {
        # オブジェクトの後始末
        $dialog.Dispose()
        return
    }
    #[System.Windows.Forms.MessageBox]::Show($dialog.FileName)
    $textBox2.Text = Split-Path -Parent $dialog.FileName
}
$browseButton2.Add_Click($Browse2)

$label_format = New-Object System.Windows.Forms.Label
$label_format.Location = New-Object System.Drawing.Point(20,210)
$label_format.Size = New-Object System.Drawing.Size(100,18)
$label_format.Text = '画像形式'
$label_format.Font = $fontDefault
$form.Controls.Add($label_format)

# 画像形式のドロップダウンリスト（ComboBoxでDropDownListスタイルにするとドロップダウンリストになる）
$comboBox_format = New-Object System.Windows.Forms.ComboBox
$comboBox_format.Location = New-Object System.Drawing.Point(20,230)
$comboBox_format.size = New-Object System.Drawing.Size(100,30)
$comboBox_format.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$comboBox_format.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
$fontCombobox = New-Object System.Drawing.Font("Calibri",11)
$comboBox_format.Font = $fontCombobox
# ドロップダウンリストに項目を追加
[void] $comboBox_format.Items.Add('.jpg/.jpeg')
[void] $comboBox_format.Items.Add('.png')
[void] $comboBox_format.Items.Add('.bmp')
[void] $comboBox_format.Items.Add('.gif')
[void] $comboBox_format.Items.Add('.tif/.tiff')
$comboBox_format.SelectedIndex = 0  # コンボボックスの既定の値
$form.Controls.Add($comboBox_format)

# サイズ指定方法のラジオボタン等
$groupbox_units = New-Object System.Windows.Forms.GroupBox
$groupbox_units.Location = New-Object System.Drawing.Point(200,210)
$groupbox_units.Size = New-Object System.Drawing.Size(190,80)
$groupbox_units.Text = '縮小サイズ'
$groupbox_units.Font = $fontDefault

$radiobutton_pixel = New-Object System.Windows.Forms.RadioButton
$radiobutton_pixel.Location = New-Object System.Drawing.Point(10,20)  # GroupBoxにAddするときはGroupBox内での相対座標
$radiobutton_pixel.Size = New-Object System.Drawing.Size(20,20)
$radiobutton_pixel.Checked = $True
#$radiobutton_pixel.Text = 'ピクセル'

$textBox_pixel = New-Object System.Windows.Forms.TextBox
$textBox_pixel.Location = New-Object System.Drawing.Point(40,20)
$textBox_pixel.Size = New-Object System.Drawing.Size(60,18)
$textBox_pixel.Font =  $fontTextbox
$textBox_pixel.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Right

$label_pixel = New-Object System.Windows.Forms.Label
$label_pixel.Location = New-Object System.Drawing.Point(110,20)
$label_pixel.Size = New-Object System.Drawing.Size(60,18)
$label_pixel.Text = 'ピクセル'

$radiobutton_ratio = New-Object System.Windows.Forms.RadioButton
$radiobutton_ratio.Location = New-Object System.Drawing.Point(10,50)
$radiobutton_ratio.Size = New-Object System.Drawing.Size(20,20)
#$radiobutton_ratio.Text = '%'

$textBox_ratio = New-Object System.Windows.Forms.TextBox
$textBox_ratio.Location = New-Object System.Drawing.Point(40,50)
$textBox_ratio.Size = New-Object System.Drawing.Size(60,18)
$textBox_ratio.Font =  $fontTextbox
$textBox_ratio.TextAlign = [System.Windows.Forms.HorizontalAlignment]::Right

$label_ratio = New-Object System.Windows.Forms.Label
$label_ratio.Location = New-Object System.Drawing.Point(110,50)
$label_ratio.Size = New-Object System.Drawing.Size(60,18)
$label_ratio.Text = '%'

# グループにラジオボタンを入れる
$groupbox_units.Controls.AddRange(@($radiobutton_pixel,$radiobutton_ratio,$textBox_pixel,$label_pixel,$textBox_ratio,$label_ratio))
#$groupbox_units.Controls.Add($radiobutton_pixel)
$form.Controls.Add($groupbox_units)
#$form.Controls.AddRange(@($groupbox_units))

$label_submit = New-Object System.Windows.Forms.Label
$label_submit.Location = New-Object System.Drawing.Point(280,305)
$label_submit.Size = New-Object System.Drawing.Size(70,20)
$label_submit.Text = '縮小処理'
$label_submit.Font = $fontDefault
$form.Controls.Add($label_submit)

$submitButton = New-Object System.Windows.Forms.Button
$submitButton.Location = New-Object System.Drawing.Point(360,300)
$submitButton.Size = New-Object System.Drawing.Size(90,23)
$submitButton.Text = '実行'
$submitButton.Font = $fontDefault
#$submitButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
#$form.AcceptButton = $okButton
$form.Controls.Add($submitButton)

# submitボタンのクリックイベント
$Submit = {

    $inputFolder = $textBox1.Text
    $outputFolder = $textBox2.Text
    # 入力フォルダー、出力フォルダーの存在チェック
    if ($inputFolder -eq "" -or -not (Test-Path $inputFolder)) {
        $msgBoxResult = $msgBox::Show($dummyWindow,'元画像のフォルダーが正しくありません', 'ReducePic', $buttonsOK, $iconErr, $defaultButton)
        return
    }
    if ($outputFolder -eq "" -or -not (Test-Path $outputFolder)) {
        $msgBoxResult = $msgBox::Show($dummyWindow,'出力先画像のフォルダーが正しくありません', 'ReducePic', $buttonsOK, $iconErr, $defaultButton)
        return
    }
    if ($inputFolder -eq $outputFolder) {
        $msgBoxResult = $msgBox::Show($dummyWindow,'出力先は元画像と異なるフォルダーを指定してください', 'ReducePic', $buttonsOK, $iconErr, $defaultButton)
        return
    }
    [int] $sizePixel = 0
    [int] $sizeRatio = 0
    if ($radiobutton_pixel.Checked) {
        if ([int]::TryParse($textBox_pixel.Text, [ref]$null)) {  # 型チェック(int)
            $sizePixel = $textBox_pixel.Text
        }
        if ($sizePixel -le 0) {
            $msgBoxResult = $msgBox::Show($dummyWindow,'サイズには1以上の整数を指定してください', 'ReducePic', $buttonsOK, $iconErr, $defaultButton)
            return
        }
    } elseif ($radiobutton_ratio.Checked) {
        if ([int]::TryParse($textBox_ratio.Text, [ref]$null)) {  # 型チェック(int)
            $sizeRatio = $textBox_ratio.Text
        }
        if ($sizeRatio -lt 1 -or $sizeRatio -gt 99) {
            $msgBoxResult = $msgBox::Show($dummyWindow,'縮小割合には1～99の整数を指定してください', 'ReducePic', $buttonsOK, $iconErr, $defaultButton)
            return
        }
    }
        <#
        $msgBoxResult = $msgBox::Show($dummyWindow, $comboBox_format.SelectedIndex, 'ReducePic', $buttonsOK, $iconInfo, $defaultButton)
        $msgBoxResult = $msgBox::Show($dummyWindow, $comboBox_format.SelectedItem, 'ReducePic', $buttonsOK, $iconInfo, $defaultButton)
        $msgBoxResult = $msgBox::Show($dummyWindow, $comboBox_format.SelectedText, 'ReducePic', $buttonsOK, $iconInfo, $defaultButton)
        $msgBoxResult = $msgBox::Show($dummyWindow, $comboBox_format.SelectedValue, 'ReducePic', $buttonsOK, $iconInfo, $defaultButton)
        return
        #>

    $form.Controls.Remove($submitButton)
    $label_submit.Text = 'Wait ...'
    $form.Refresh()  # 再描画しないとラベルのテキスト変更が反映されなかった

    switch ($comboBox_format.SelectedItem) {
        '.jpg/.jpeg' {
            $picExt1 = '*.jpg'
            $picExt2 = '*.jpeg'
        }
        '.png' {
            $picExt1 = '*.png'
            $picExt2 = $picExtension1
        }
        '.bmp' {
            $picExt1 = '*.bmp'
            $picExt2 = $picExtension1
        }
        '.gif' {
            $picExt1 = '*.gif'
            $picExt2 = $picExtension1
        }
        '.tif/.tiff' {
            $picExt1 = '*.tif'
            $picExt2 = '*.tiff'
        }
    }

    $cnt = 0
	Get-ChildItem -Path $inputFolder -File | Where-Object{$_.Name -like $picExt1 -or $_.Name -like $picExt2} | ForEach-Object {
        $newFile = $outputFolder + "\" + $_.Name
        resizeImage $_.FullName $newFile $picExt1 $sizePixel $sizeRatio

        $cnt++
        $label_submit.Text = '(' + $cnt + ')'
        $form.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
    }

    $label_submit.Text = "${cnt}件 完了"
    $form.Controls.Add($submitButton)

}
$submitButton.Add_Click($Submit)

$label_config = New-Object System.Windows.Forms.Label
$label_config.Location = New-Object System.Drawing.Point(10,340)
$label_config.Size = New-Object System.Drawing.Size(400,40)
#$label_config.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$label_config.Text = ''
if ($interpolationName -ne '') {
    $label_config.Text = $label_config.Text + 'Interpolation='+ $interpolationName + "`r`n"
}
if ($jpegQuality -ne 999) {
    $label_config.Text = $label_config.Text + 'Quality='+ $jpegQuality
}
$fontDefault = New-Object System.Drawing.Font("Yu Gothic UI",$label_config.Font.Size)
$label_config.Font = $fontDefault
$form.Controls.Add($label_config)

<#
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(230,90)
$cancelButton.Size = New-Object System.Drawing.Size(60,23)
$cancelButton.Text = 'Close'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)
#>

# フォームをアクティブにしてtextBoxにフォーカスを設定する
$form.Add_Shown({$textBox1.Select()})

# [High DPI] 最後にResumeLayoutを実行してからShowDialogする
$form.ResumeLayout()

$result = $form.ShowDialog()

<#
if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
{
    # nop
}
#>
# ログ記録
<#
if ($isLog) {
    $userName = (whoami).Split('\')[1]
    writeLog "[image]$DateTime.$Format [user]$userName"
}
#>

$form.Dispose()

exit
