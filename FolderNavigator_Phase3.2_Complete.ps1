# ===============================================
# フォルダナビゲーター Phase 3.2 Complete
# Version: 3.2.0
# Date: 2024-11-23
# Author: KENJI
# ===============================================
# 
# 【機能一覧】
# - フォルダツリー表示（階層構造）
# - ファイル・フォルダ一覧表示
# - リアルタイム検索（新機能！）
# - 新規フォルダ作成（試験項目連番対応）
# - フォルダを開く（エクスプローラー起動）
# - 一括リネーム機能（実装済み）
#   - プレフィックス追加
#   - サフィックス追加
#   - 文字列置換
#   - 連番付与
#   - プレビュー機能
#   - 実際のリネーム実行
# - ドラッグ&ドロップ機能
#   - Excel/テキストファイル対応
#   - 上書き確認ダイアログ
#   - 自動リスト更新
# - 履歴機能（最近使用したフォルダ）
# - ネットワークドライブ対応
# - UI完全日本語化（新機能！）
# ===============================================

# STAモードチェック（WPF必須）
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host " WPFにはSTAモードが必要です" -ForegroundColor Yellow
    Write-Host " STAモードで再起動します..." -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    
    # STAモードで再起動（-Waitを削除して独立プロセスとして起動）
    Start-Process powershell.exe -ArgumentList @(
        "-STA",
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", "`"$PSCommandPath`""
    )
    exit
}

# 必要なアセンブリの読み込み
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# グローバル変数
$script:currentPath = ""
$script:history = @()
$script:maxHistory = 10
$script:version = "3.2.0"
$script:previewItems = @()

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " フォルダナビゲーター v$script:version 起動中..." -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# メインウィンドウXAML
$mainXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="フォルダナビゲーター v3.2 - Phase 3.2" Width="1200" Height="800"
    WindowStartupLocation="CenterScreen"
    AllowDrop="True">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- ヘッダー -->
        <Border Grid.Row="0" Background="#34495e" Padding="10">
            <StackPanel>
                <TextBlock Text="フォルダナビゲーター + 一括リネーム + ドラッグ&amp;ドロップ" FontSize="18" FontWeight="Bold" 
                          Foreground="White" Margin="0,0,0,10"/>
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="フォルダパス：" Foreground="White" 
                              VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <TextBox Name="txtPath" Grid.Column="1" FontSize="12" Padding="5"/>
                    <Button Name="btnBrowse" Grid.Column="2" Content="参照" Padding="10,5" 
                           Margin="5" MinWidth="80"/>
                    <Button Name="btnLoad" Grid.Column="3" Content="読込" Padding="10,5" 
                           Margin="5" MinWidth="80"/>
                </Grid>
            </StackPanel>
        </Border>
        
        <!-- 検索バー -->
        <Border Grid.Row="1" Background="#ecf0f1" Padding="10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="300"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="検索（配下すべて）：" VerticalAlignment="Center" Margin="0,0,10,0"/>
                <TextBox Name="txtSearch" Grid.Column="1" FontSize="12" Padding="5" 
                         ToolTip="配下のすべてのフォルダから検索します"/>
                <Button Name="btnSearchClear" Grid.Column="2" Content="×" Width="30" Margin="5,0,0,0" 
                        ToolTip="クリア" FontSize="16" FontWeight="Bold"/>
                <TextBlock Grid.Column="3" Text="履歴：" VerticalAlignment="Center" Margin="20,0,10,0"/>
                <ComboBox Name="cmbHistory" Grid.Column="4" MaxWidth="400" HorizontalAlignment="Left"/>
            </Grid>
        </Border>
        
        <!-- メインコンテンツ -->
        <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="2*" MinWidth="300"/>
                <ColumnDefinition Width="5"/>
                <ColumnDefinition Width="3*" MinWidth="400"/>
            </Grid.ColumnDefinitions>
            
            <!-- 左パネル: フォルダツリー -->
            <Border Grid.Column="0" BorderBrush="#bdc3c7" BorderThickness="1" Margin="5">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <TextBlock Grid.Row="0" Text="フォルダツリー" FontSize="14" FontWeight="Bold" 
                              Padding="10,5" Background="#2c3e50" Foreground="White"/>
                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
                        <TreeView Name="treeView" FontSize="12" Padding="5" BorderThickness="0"/>
                    </ScrollViewer>
                </Grid>
            </Border>
            
            <!-- スプリッター -->
            <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Stretch" 
                         VerticalAlignment="Stretch" Background="#95a5a6"/>
            
            <!-- 右パネル: ファイル一覧 -->
            <Border Grid.Column="2" BorderBrush="#bdc3c7" BorderThickness="1" Margin="5">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <TextBlock Grid.Row="0" Text="ファイル一覧（Excel/テキストをドロップ）" FontSize="14" FontWeight="Bold" 
                              Padding="10,5" Background="#2c3e50" Foreground="White"/>
                    <DataGrid Name="dataGrid" Grid.Row="1" AutoGenerateColumns="False" 
                             CanUserAddRows="False" GridLinesVisibility="None" 
                             AlternatingRowBackground="#f8f9fa"
                             AllowDrop="True"
                             VerticalScrollBarVisibility="Auto"
                             HorizontalScrollBarVisibility="Auto"
                             EnableRowVirtualization="True"
                             EnableColumnVirtualization="True">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="種類" Binding="{Binding Type}" Width="80"/>
                            <DataGridTextColumn Header="名前" Binding="{Binding Name}" Width="2*"/>
                            <DataGridTextColumn Header="サイズ" Binding="{Binding Size}" Width="100"/>
                            <DataGridTextColumn Header="更新日時" Binding="{Binding Modified}" Width="150"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </Border>
        </Grid>
        
        <!-- フッター -->
        <Border Grid.Row="3" Background="#ecf0f1" Padding="10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Name="txtStatus" Grid.Column="0" Text="準備完了 - ファイルをドラッグ＆ドロップしてください" VerticalAlignment="Center"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <Button Name="btnOpenFolder" Content="フォルダを開く" Padding="10,5" Margin="5" MinWidth="100"/>
                    <Button Name="btnNewFolder" Content="新規フォルダ" Padding="10,5" Margin="5" MinWidth="100"/>
                    <Button Name="btnTestFolder" Content="試験フォルダ" Padding="10,5" Margin="5" MinWidth="100"
                           ToolTip="試験項目フォルダを自動作成"/>
                    <Button Name="btnBatchRename" Content="一括リネーム" Padding="10,5" Margin="5" MinWidth="100" 
                           Background="#e74c3c" Foreground="White" FontWeight="Bold"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# FileItemクラス
Add-Type @"
public class FileItem {
    public string Type { get; set; }
    public string Name { get; set; }
    public string Size { get; set; }
    public string Modified { get; set; }
    public string FullPath { get; set; }
}
"@ -ErrorAction SilentlyContinue

# メインウィンドウ作成
[xml]$x = $mainXaml
$reader = New-Object System.Xml.XmlNodeReader $x
$ErrorActionPreference = 'SilentlyContinue'
$mainWindow = [Windows.Markup.XamlReader]::Load($reader)
$ErrorActionPreference = 'Continue'

# コントロール取得
$txtPath = $mainWindow.FindName("txtPath")
$btnBrowse = $mainWindow.FindName("btnBrowse")
$btnLoad = $mainWindow.FindName("btnLoad")
$cmbHistory = $mainWindow.FindName("cmbHistory")
$txtSearch = $mainWindow.FindName("txtSearch")
$btnSearchClear = $mainWindow.FindName("btnSearchClear")
$treeView = $mainWindow.FindName("treeView")
$dataGrid = $mainWindow.FindName("dataGrid")
$txtStatus = $mainWindow.FindName("txtStatus")
$btnOpenFolder = $mainWindow.FindName("btnOpenFolder")
$btnNewFolder = $mainWindow.FindName("btnNewFolder")
$btnTestFolder = $mainWindow.FindName("btnTestFolder")
$btnBatchRename = $mainWindow.FindName("btnBatchRename")

# ===============================================
# ドラッグ&ドロップ機能
# ===============================================

# サポートする拡張子
$script:supportedExtensions = @('.xlsx', '.xls', '.xlsm', '.txt', '.csv')

# ドラッグエンター（視覚フィードバック）
$mainWindow.Add_DragEnter({
    param($sender, $e)
    
    if ($e.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $files = $e.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
        $validFiles = $files | Where-Object {
            $ext = [System.IO.Path]::GetExtension($_).ToLower()
            $script:supportedExtensions -contains $ext
        }
        
        if ($validFiles.Count -gt 0) {
            $e.Effects = [System.Windows.DragDropEffects]::Copy
            $txtStatus.Text = "ファイルをドロップしてコピー..."
        }
        else {
            $e.Effects = [System.Windows.DragDropEffects]::None
            $txtStatus.Text = "サポートされていないファイル形式です"
        }
    }
    $e.Handled = $true
})

# ドラッグリーブ
$mainWindow.Add_DragLeave({
    $txtStatus.Text = "準備完了 - ファイルをドラッグ＆ドロップしてください"
})

# ドロップ（ファイルコピー処理）
$mainWindow.Add_Drop({
    param($sender, $e)
    
    if (!$script:currentPath) {
        [System.Windows.MessageBox]::Show(
            "先にフォルダを選択してください",
            "エラー",
            "OK",
            "Warning"
        )
        return
    }
    
    if ($e.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $files = $e.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
        $copiedCount = 0
        $skippedCount = 0
        
        foreach ($file in $files) {
            $ext = [System.IO.Path]::GetExtension($file).ToLower()
            
            # 拡張子チェック
            if ($script:supportedExtensions -contains $ext) {
                $fileName = [System.IO.Path]::GetFileName($file)
                $destPath = Join-Path $script:currentPath $fileName
                
                # 同名ファイルチェック
                if (Test-Path $destPath) {
                    $result = [System.Windows.MessageBox]::Show(
                        "ファイル '$fileName' は既に存在します。`n上書きしますか？",
                        "上書き確認",
                        "YesNo",
                        "Question"
                    )
                    
                    if ($result -eq "No") {
                        $skippedCount++
                        continue
                    }
                }
                
                try {
                    Copy-Item -Path $file -Destination $destPath -Force
                    $copiedCount++
                    Write-Host "コピー完了: $fileName" -ForegroundColor Green
                }
                catch {
                    [System.Windows.MessageBox]::Show(
                        "ファイルのコピーに失敗しました:`n$fileName`n`nエラー: $_",
                        "エラー",
                        "OK",
                        "Error"
                    )
                }
            }
        }
        
        # 結果表示
        if ($copiedCount -gt 0) {
            $message = "コピー完了: $copiedCount 件"
            if ($skippedCount -gt 0) {
                $message += " (スキップ: $skippedCount 件)"
            }
            $txtStatus.Text = $message
            
            # ファイルリスト更新
            Load-FileList $script:currentPath
        }
        else {
            $txtStatus.Text = "コピーされたファイルはありません"
        }
    }
    
    $e.Handled = $true
})

# DataGrid用ドラッグエンター
$dataGrid.Add_DragEnter({
    param($sender, $e)
    
    if ($e.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effects = [System.Windows.DragDropEffects]::Copy
    }
    $e.Handled = $true
})

# DataGrid用ドロップ
$dataGrid.Add_Drop({
    param($sender, $e)
    
    # メインウィンドウのドロップ処理を呼び出し
    $mainWindow_Drop = $mainWindow | Get-Member -Name "Drop" -MemberType Event
    $mainWindow.RaiseEvent($e)
})

# ===============================================
# 一括リネーム機能
# ===============================================
function Show-BatchRenameDialog {
    param($currentPath)
    
    if (!$currentPath) {
        [System.Windows.Forms.MessageBox]::Show(
            "フォルダを選択してください", 
            "エラー", "OK", "Warning")
        return
    }
    
    # リネームフォーム作成
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "一括リネームツール"
    $form.Size = New-Object System.Drawing.Size(750, 650)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("メイリオ", 10)
    
    # 説明パネル
    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Location = New-Object System.Drawing.Point(20, 10)
    $lblInfo.Size = New-Object System.Drawing.Size(700, 60)
    $lblInfo.Text = "ファイル・フォルダの名前を一括で変更します`n現在のフォルダ: $currentPath"
    $lblInfo.BackColor = [System.Drawing.Color]::LightBlue
    $lblInfo.Padding = New-Object System.Windows.Forms.Padding(10)
    
    # リネーム方式
    $lblType = New-Object System.Windows.Forms.Label
    $lblType.Location = New-Object System.Drawing.Point(20, 90)
    $lblType.Size = New-Object System.Drawing.Size(120, 25)
    $lblType.Text = "リネーム方式:"
    
    $cmbType = New-Object System.Windows.Forms.ComboBox
    $cmbType.Location = New-Object System.Drawing.Point(150, 87)
    $cmbType.Size = New-Object System.Drawing.Size(300, 30)
    $cmbType.DropDownStyle = "DropDownList"
    $cmbType.Items.AddRange(@(
        "先頭に文字追加（プレフィックス）",
        "末尾に文字追加（サフィックス）",
        "文字列を置換",
        "連番を追加"
    ))
    $cmbType.SelectedIndex = 0
    
    # 使い方
    $grpUsage = New-Object System.Windows.Forms.GroupBox
    $grpUsage.Location = New-Object System.Drawing.Point(470, 85)
    $grpUsage.Size = New-Object System.Drawing.Size(250, 120)
    $grpUsage.Text = "【使い方】"
    
    $lblUsage = New-Object System.Windows.Forms.Label
    $lblUsage.Location = New-Object System.Drawing.Point(10, 20)
    $lblUsage.Size = New-Object System.Drawing.Size(230, 90)
    $lblUsage.Text = "1. リネーム方式を選択`n2. 文字を入力`n3. プレビューで確認`n4. 実行ボタンでリネーム"
    
    $grpUsage.Controls.Add($lblUsage)
    
    # 入力欄1
    $lblInput1 = New-Object System.Windows.Forms.Label
    $lblInput1.Location = New-Object System.Drawing.Point(20, 130)
    $lblInput1.Size = New-Object System.Drawing.Size(120, 25)
    $lblInput1.Text = "追加する文字:"
    
    $txtInput1 = New-Object System.Windows.Forms.TextBox
    $txtInput1.Location = New-Object System.Drawing.Point(150, 127)
    $txtInput1.Size = New-Object System.Drawing.Size(300, 30)
    $txtInput1.Text = "2024_"
    
    # 入力欄2（置換用）
    $lblInput2 = New-Object System.Windows.Forms.Label
    $lblInput2.Location = New-Object System.Drawing.Point(20, 165)
    $lblInput2.Size = New-Object System.Drawing.Size(120, 25)
    $lblInput2.Text = "置換後の文字:"
    $lblInput2.Visible = $false
    
    $txtInput2 = New-Object System.Windows.Forms.TextBox
    $txtInput2.Location = New-Object System.Drawing.Point(150, 162)
    $txtInput2.Size = New-Object System.Drawing.Size(300, 30)
    $txtInput2.Visible = $false
    
    # プレビューボタン
    $btnPreview = New-Object System.Windows.Forms.Button
    $btnPreview.Location = New-Object System.Drawing.Point(150, 210)
    $btnPreview.Size = New-Object System.Drawing.Size(150, 40)
    $btnPreview.Text = "プレビュー"
    $btnPreview.BackColor = [System.Drawing.Color]::LightGreen
    $btnPreview.Font = New-Object System.Drawing.Font("メイリオ", 11, [System.Drawing.FontStyle]::Bold)
    
    # プレビューラベル
    $lblPreview = New-Object System.Windows.Forms.Label
    $lblPreview.Location = New-Object System.Drawing.Point(20, 265)
    $lblPreview.Size = New-Object System.Drawing.Size(700, 25)
    $lblPreview.Text = "変更プレビュー（変更前 → 変更後）:"
    $lblPreview.Font = New-Object System.Drawing.Font("メイリオ", 10, [System.Drawing.FontStyle]::Bold)
    
    # プレビューリスト
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(20, 295)
    $listBox.Size = New-Object System.Drawing.Size(700, 250)
    $listBox.Font = New-Object System.Drawing.Font("MS Gothic", 10)
    $listBox.HorizontalScrollbar = $true
    
    # 実行・キャンセル
    $btnExecute = New-Object System.Windows.Forms.Button
    $btnExecute.Location = New-Object System.Drawing.Point(520, 560)
    $btnExecute.Size = New-Object System.Drawing.Size(100, 40)
    $btnExecute.Text = "実行"
    $btnExecute.BackColor = [System.Drawing.Color]::Tomato
    $btnExecute.ForeColor = [System.Drawing.Color]::White
    $btnExecute.Font = New-Object System.Drawing.Font("メイリオ", 11, [System.Drawing.FontStyle]::Bold)
    $btnExecute.Enabled = $false
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(630, 560)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 40)
    $btnCancel.Text = "キャンセル"
    
    # イベント処理
    $cmbType.Add_SelectedIndexChanged({
        switch ($cmbType.SelectedIndex) {
            0 { # プレフィックス
                $lblInput1.Text = "追加する文字:"
                $txtInput1.Text = "2024_"
                $lblInput2.Visible = $false
                $txtInput2.Visible = $false
                $lblUsage.Text = "ファイル名の先頭に`n文字を追加します`n`n例: test.txt`n→ 2024_test.txt"
            }
            1 { # サフィックス
                $lblInput1.Text = "追加する文字:"
                $txtInput1.Text = "_完了"
                $lblInput2.Visible = $false
                $txtInput2.Visible = $false
                $lblUsage.Text = "ファイル名の末尾に`n文字を追加します`n`n例: test.txt`n→ test_完了.txt"
            }
            2 { # 置換
                $lblInput1.Text = "置換前の文字:"
                $txtInput1.Text = "2025"
                $lblInput2.Visible = $true
                $txtInput2.Visible = $true
                $txtInput2.Text = "2024"
                $lblUsage.Text = "文字列を置換します`n`n例: 2025_test.txt`n→ 2024_test.txt"
            }
            3 { # 連番
                $lblInput1.Text = "プレフィックス:"
                $txtInput1.Text = "File_"
                $lblInput2.Visible = $false
                $txtInput2.Visible = $false
                $lblUsage.Text = "連番を追加します`n`n例: test.txt`n→ File_001_test.txt"
            }
        }
    })
    
    # プレビュー処理
    $btnPreview.Add_Click({
        $listBox.Items.Clear()
        $script:previewItems = @()  # 初期化
        $items = Get-ChildItem -Path $currentPath -ErrorAction SilentlyContinue | Select-Object -First 30
        $counter = 1
        
        foreach ($item in $items) {
            $newName = ""
            $type = if ($item.PSIsContainer) { "[フォルダ]" } else { "[ファイル]" }
            
            switch ($cmbType.SelectedIndex) {
                0 { # プレフィックス
                    $newName = $txtInput1.Text + $item.Name
                }
                1 { # サフィックス
                    if ($item.PSIsContainer) {
                        $newName = $item.Name + $txtInput1.Text
                    } else {
                        $name = [System.IO.Path]::GetFileNameWithoutExtension($item.Name)
                        $ext = [System.IO.Path]::GetExtension($item.Name)
                        $newName = $name + $txtInput1.Text + $ext
                    }
                }
                2 { # 置換
                    if ($txtInput2.Visible) {
                        $newName = $item.Name.Replace($txtInput1.Text, $txtInput2.Text)
                    } else {
                        $newName = $item.Name
                    }
                }
                3 { # 連番
                    $newName = $txtInput1.Text + "{0:D3}_" -f $counter + $item.Name
                    $counter++
                }
            }
            
            # プレビューデータを保存
            $script:previewItems += @{
                OldPath = $item.FullName
                OldName = $item.Name
                NewName = $newName
                IsFolder = $item.PSIsContainer
            }
            
            $listBox.Items.Add("$type $($item.Name) → $newName")
        }
        
        $btnExecute.Enabled = $listBox.Items.Count -gt 0
    })
    
    # 実行処理（実際のリネーム）
    $btnExecute.Add_Click({
        $result = [System.Windows.Forms.MessageBox]::Show(
            "本当にリネームを実行しますか？`n`nこの操作は取り消せません。",
            "最終確認",
            "YesNo",
            "Warning"
        )
        
        if ($result -eq "Yes") {
            $successCount = 0
            $failCount = 0
            $errorMessages = @()
            
            try {
                foreach ($previewItem in $script:previewItems) {
                    $oldPath = $previewItem.OldPath
                    $newName = $previewItem.NewName
                    $parentPath = Split-Path $oldPath -Parent
                    $newPath = Join-Path $parentPath $newName
                    
                    # 同名チェック
                    if ($oldPath -ne $newPath -and (Test-Path $newPath)) {
                        $errorMsg = "同名のファイル/フォルダが存在: $newName"
                        $errorMessages += $errorMsg
                        [System.Windows.Forms.MessageBox]::Show(
                            $errorMsg + "`n`n処理を中止します。",
                            "エラー",
                            "OK",
                            "Error"
                        )
                        $failCount++
                        break
                    }
                    
                    # リネーム実行
                    if ($oldPath -ne $newPath) {
                        Rename-Item -Path $oldPath -NewName $newName -ErrorAction Stop
                        $successCount++
                    }
                }
                
                # 結果表示
                if ($failCount -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "リネームが完了しました！`n`n成功: $successCount 件",
                        "完了",
                        "OK",
                        "Information"
                    )
                    
                    # ファイルリストを更新
                    Load-FileList $currentPath
                    $form.Close()
                }
            }
            catch {
                $errorMsg = "リネーム中にエラーが発生しました: $_"
                $errorMessages += $errorMsg
                [System.Windows.Forms.MessageBox]::Show(
                    $errorMsg + "`n`n成功: $successCount 件`n失敗: $($failCount + 1) 件",
                    "エラー",
                    "OK",
                    "Error"
                )
            }
        }
    })
    
    $btnCancel.Add_Click({
        $form.Close()
    })
    
    # コントロール追加
    $form.Controls.AddRange(@(
        $lblInfo,
        $lblType, $cmbType,
        $grpUsage,
        $lblInput1, $txtInput1,
        $lblInput2, $txtInput2,
        $btnPreview,
        $lblPreview, $listBox,
        $btnExecute, $btnCancel
    ))
    
    # フォーム表示
    $form.ShowDialog() | Out-Null
}

# ===============================================
# メイン機能関数
# ===============================================

# フォルダツリー読み込み
function Load-FolderTree {
    param($path)
    
    if (!(Test-Path $path)) {
        [System.Windows.MessageBox]::Show(
            "指定されたパスが見つかりません:`n$path", 
            "エラー", "OK", "Error")
        return
    }
    
    $txtStatus.Text = "読み込み中..."
    $treeView.Items.Clear()
    
    try {
        $rootItem = New-Object System.Windows.Controls.TreeViewItem
        $rootItem.Header = Split-Path $path -Leaf
        if (!$rootItem.Header) { $rootItem.Header = $path }
        $rootItem.Tag = $path
        $rootItem.IsExpanded = $true
        
        Load-SubFolders -parentItem $rootItem -path $path -depth 0 -maxDepth 2
        
        $treeView.Items.Add($rootItem)
        
        $script:currentPath = $path
        $txtStatus.Text = "現在: $path"
        
        # 履歴に追加
        if ($script:history -notcontains $path) {
            $script:history = @($path) + $script:history | Select-Object -First $script:maxHistory
            Update-HistoryComboBox
        }
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "フォルダの読み込みに失敗しました:`n$_", 
            "エラー", "OK", "Error")
        $txtStatus.Text = "エラーが発生しました"
    }
}

# サブフォルダ読み込み
function Load-SubFolders {
    param($parentItem, $path, $depth, $maxDepth)
    
    if ($depth -ge $maxDepth) { return }
    
    try {
        $folders = Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue | 
                   Where-Object { !$_.Attributes.HasFlag([System.IO.FileAttributes]::Hidden) -and 
                                 !$_.Attributes.HasFlag([System.IO.FileAttributes]::System) }
        
        foreach ($folder in $folders) {
            $item = New-Object System.Windows.Controls.TreeViewItem
            $item.Header = $folder.Name
            $item.Tag = $folder.FullName
            
            # 遅延読み込み用ダミー
            if ((Get-ChildItem -Path $folder.FullName -Directory -ErrorAction SilentlyContinue | Select-Object -First 1)) {
                $dummy = New-Object System.Windows.Controls.TreeViewItem
                $dummy.Header = "読み込み中..."
                $item.Items.Add($dummy)
            }
            
            $parentItem.Items.Add($item)
        }
    }
    catch {
        Write-Host "サブフォルダ読み込みエラー: $_" -ForegroundColor Red
    }
}

# ファイルリスト読み込み
function Load-FileList {
    param($path)
    
    if (!(Test-Path $path)) { return }
    
    $dataGrid.Items.Clear()
    
    try {
        # フォルダを追加
        $folders = Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue
        foreach ($folder in $folders) {
            $item = New-Object FileItem
            $item.Type = "フォルダ"
            $item.Name = $folder.Name
            $item.Size = "-"
            $item.Modified = $folder.LastWriteTime.ToString("yyyy/MM/dd HH:mm")
            $item.FullPath = $folder.FullName
            $dataGrid.Items.Add($item)
        }
        
        # ファイルを追加
        $files = Get-ChildItem -Path $path -File -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            $item = New-Object FileItem
            $item.Type = "ファイル"
            $item.Name = $file.Name
            
            # サイズフォーマット
            if ($file.Length -lt 1KB) {
                $item.Size = "{0} B" -f $file.Length
            }
            elseif ($file.Length -lt 1MB) {
                $item.Size = "{0:N1} KB" -f ($file.Length / 1KB)
            }
            elseif ($file.Length -lt 1GB) {
                $item.Size = "{0:N1} MB" -f ($file.Length / 1MB)
            }
            else {
                $item.Size = "{0:N1} GB" -f ($file.Length / 1GB)
            }
            
            $item.Modified = $file.LastWriteTime.ToString("yyyy/MM/dd HH:mm")
            $item.FullPath = $file.FullName
            $dataGrid.Items.Add($item)
        }
        
        $txtStatus.Text = "現在: $path (フォルダ: $($folders.Count) / ファイル: $($files.Count))"
    }
    catch {
        Write-Host "ファイルリスト読み込みエラー: $_" -ForegroundColor Red
    }
}

# 履歴更新
function Update-HistoryComboBox {
    $cmbHistory.Items.Clear()
    foreach ($item in $script:history) {
        $cmbHistory.Items.Add($item)
    }
}

# 新規フォルダ作成
function Create-NewFolder {
    if (!$script:currentPath) {
        [System.Windows.MessageBox]::Show(
            "先にフォルダを選択してください", 
            "情報", "OK", "Information")
        return
    }
    
    $folderName = [Microsoft.VisualBasic.Interaction]::InputBox(
        "新しいフォルダ名を入力してください:", 
        "新規フォルダ作成", "")
    
    if ($folderName -eq "") { return }
    
    $newPath = Join-Path $script:currentPath $folderName
    
    try {
        if (Test-Path $newPath) {
            [System.Windows.MessageBox]::Show(
                "同名のフォルダが既に存在します", 
                "エラー", "OK", "Warning")
            return
        }
        
        New-Item -ItemType Directory -Path $newPath | Out-Null
        [System.Windows.MessageBox]::Show(
            "フォルダを作成しました:`n$folderName", 
            "成功", "OK", "Information")
        
        Load-FolderTree $txtPath.Text
        Load-FileList $script:currentPath
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "フォルダの作成に失敗しました:`n$_", 
            "エラー", "OK", "Error")
    }
}

# 試験項目フォルダ作成
function Create-TestFolder {
    if (!$script:currentPath) {
        [System.Windows.MessageBox]::Show(
            "先にフォルダを選択してください", 
            "情報", "OK", "Information")
        return
    }
    
    # 連番生成
    for ($i = 1; $i -le 999; $i++) {
        $folderName = "試験項目{0:D2}" -f $i
        $newPath = Join-Path $script:currentPath $folderName
        if (!(Test-Path $newPath)) {
            try {
                New-Item -ItemType Directory -Path $newPath | Out-Null
                [System.Windows.MessageBox]::Show(
                    "試験項目フォルダを作成しました:`n$folderName", 
                    "成功", "OK", "Information")
                
                Load-FolderTree $txtPath.Text
                Load-FileList $script:currentPath
                break
            }
            catch {
                [System.Windows.MessageBox]::Show(
                    "フォルダの作成に失敗しました:`n$_", 
                    "エラー", "OK", "Error")
                break
            }
        }
    }
}

# ===============================================
# イベントハンドラ
# ===============================================

# 参照ボタン
$btnBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "フォルダを選択してください"
    $dialog.ShowNewFolderButton = $true
    
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtPath.Text = $dialog.SelectedPath
        Load-FolderTree $dialog.SelectedPath
    }
})

# 読み込みボタン
$btnLoad.Add_Click({
    $path = $txtPath.Text.Trim()
    if ($path) {
        Load-FolderTree $path
    }
})

# Enterキーで読み込み
$txtPath.Add_KeyDown({
    if ($_.Key -eq "Return") {
        $path = $txtPath.Text.Trim()
        if ($path) {
            Load-FolderTree $path
        }
    }
})

# 履歴選択
$cmbHistory.Add_SelectionChanged({
    if ($cmbHistory.SelectedItem) {
        $txtPath.Text = $cmbHistory.SelectedItem
        Load-FolderTree $cmbHistory.SelectedItem
    }
})

# ツリービュー選択
$treeView.Add_SelectedItemChanged({
    if ($_.NewValue) {
        $selectedPath = $_.NewValue.Tag
        $script:currentPath = $selectedPath
        Load-FileList $selectedPath
        
        # 遅延読み込み
        if ($_.NewValue.Items.Count -eq 1 -and $_.NewValue.Items[0].Header -eq "読み込み中...") {
            $_.NewValue.Items.Clear()
            Load-SubFolders -parentItem $_.NewValue -path $selectedPath -depth 0 -maxDepth 1
        }
    }
})

# ダブルクリック処理
$dataGrid.Add_MouseDoubleClick({
    if ($dataGrid.SelectedItem) {
        $selectedItem = $dataGrid.SelectedItem
        if ($selectedItem.Type -eq "フォルダ") {
            $txtPath.Text = $selectedItem.FullPath
            Load-FolderTree $selectedItem.FullPath
        }
        else {
            Start-Process $selectedItem.FullPath
        }
    }
})

# 検索機能（リアルタイムフィルタリング - 再帰的）
$txtSearch.Add_TextChanged({
    try {
        $searchText = $txtSearch.Text.Trim().ToLower()
        $dataGrid.Items.Clear()
        
        if (!$script:currentPath -or !(Test-Path $script:currentPath)) {
            return
        }
        
        # 検索テキストがある場合は再帰的に検索
        if (![string]::IsNullOrWhiteSpace($searchText)) {
            $txtStatus.Text = "検索中..."
            
            # 配下のすべてのフォルダを再帰的に取得
            $folders = Get-ChildItem -Path $script:currentPath -Directory -Recurse -ErrorAction SilentlyContinue |
                       Where-Object { $_.Name.ToLower().Contains($searchText) } |
                       Sort-Object FullName
            
            # 配下のすべてのファイルを再帰的に取得
            $files = Get-ChildItem -Path $script:currentPath -File -Recurse -ErrorAction SilentlyContinue |
                     Where-Object { $_.Name.ToLower().Contains($searchText) } |
                     Sort-Object FullName
            
            # フォルダ追加
            foreach ($folder in $folders) {
                $item = New-Object FileItem
                $item.Type = "フォルダ"
                # 相対パスを表示
                $relativePath = $folder.FullName.Replace($script:currentPath, "").TrimStart('\')
                $item.Name = $relativePath
                $item.Size = "-"
                $item.Modified = $folder.LastWriteTime.ToString("yyyy/MM/dd HH:mm")
                $item.FullPath = $folder.FullName
                $dataGrid.Items.Add($item)
            }
            
            # ファイル追加
            foreach ($file in $files) {
                $item = New-Object FileItem
                $item.Type = "ファイル"
                # 相対パスを表示
                $relativePath = $file.FullName.Replace($script:currentPath, "").TrimStart('\')
                $item.Name = $relativePath
                
                if ($file.Length -lt 1KB) {
                    $item.Size = "{0} B" -f $file.Length
                }
                elseif ($file.Length -lt 1MB) {
                    $item.Size = "{0:N1} KB" -f ($file.Length / 1KB)
                }
                elseif ($file.Length -lt 1GB) {
                    $item.Size = "{0:N1} MB" -f ($file.Length / 1MB)
                }
                else {
                    $item.Size = "{0:N1} GB" -f ($file.Length / 1GB)
                }
                
                $item.Modified = $file.LastWriteTime.ToString("yyyy/MM/dd HH:mm")
                $item.FullPath = $file.FullName
                $dataGrid.Items.Add($item)
            }
            
            # ステータス更新
            $totalCount = $folders.Count + $files.Count
            $txtStatus.Text = "検索結果（再帰的）: $totalCount 件（フォルダ: $($folders.Count) / ファイル: $($files.Count)）"
        }
        else {
            # 検索テキストが空の場合は現在のフォルダのみ表示
            Load-FileList $script:currentPath
        }
    }
    catch {
        Write-Host "検索エラー: $_" -ForegroundColor Red
        $txtStatus.Text = "検索エラーが発生しました"
    }
})

# 検索クリアボタン
$btnSearchClear.Add_Click({
    $txtSearch.Text = ""
    if ($script:currentPath) {
        Load-FileList $script:currentPath
    }
})

# フォルダを開くボタン
$btnOpenFolder.Add_Click({
    if ($script:currentPath -and (Test-Path $script:currentPath)) {
        Start-Process explorer.exe $script:currentPath
    }
    else {
        [System.Windows.MessageBox]::Show(
            "フォルダを選択してください", 
            "情報", "OK", "Information")
    }
})

# 新規フォルダボタン
$btnNewFolder.Add_Click({
    Create-NewFolder
})

# 試験項目フォルダボタン
$btnTestFolder.Add_Click({
    Create-TestFolder
})

# 一括リネームボタン
$btnBatchRename.Add_Click({
    Write-Host "一括リネーム機能を起動..." -ForegroundColor Yellow
    Show-BatchRenameDialog $script:currentPath
})

# ===============================================
# 初期化
# ===============================================

# デフォルトパス設定
$txtPath.Text = $env:USERPROFILE
Load-FolderTree $env:USERPROFILE

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host " 起動完了！" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "【機能一覧】" -ForegroundColor Cyan
Write-Host "  ✅ フォルダツリー表示" -ForegroundColor White
Write-Host "  ✅ ファイル・フォルダ一覧" -ForegroundColor White
Write-Host "  ✅ 新規フォルダ作成" -ForegroundColor White
Write-Host "  ✅ 試験項目フォルダ（連番）" -ForegroundColor White
Write-Host "  ✅ 一括リネーム機能（実装済み）" -ForegroundColor White
Write-Host "  ✅ ドラッグ&ドロップ（Excel/Text）" -ForegroundColor White
Write-Host "  ✅ 再帰検索機能（配下すべて）" -ForegroundColor Green
Write-Host "  ✅ 履歴機能" -ForegroundColor White
Write-Host ""

# メインウィンドウ表示
$mainWindow.ShowDialog() | Out-Null
Write-Host "アプリケーションを終了しました" -ForegroundColor Yellow

