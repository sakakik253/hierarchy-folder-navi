# ===============================================
# フォルダナビゲーター Phase 3 完全版
# Version: 3.0.0
# Date: 2024-11-23
# Author: KENJI
# ===============================================
# 
# 【機能一覧】
# - フォルダツリー表示（階層構造）
# - ファイル・フォルダ一覧表示
# - リアルタイム検索
# - 新規フォルダ作成（試験項目連番対応）
# - フォルダを開く（エクスプローラー起動）
# - 一括リネーム機能（実運用版）
#   - プレフィックス追加
#   - サフィックス追加
#   - 文字列置換
#   - 連番付与
#   - プレビュー機能
#   - エラー時即座に中止
#   - ログファイル出力
# - ドラッグ&ドロップ機能
#   - Excel/テキストファイルをコピー
# - 履歴機能（最近使用したフォルダ）
# - ネットワークドライブ対応
# ===============================================

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
$script:version = "3.0.0"
$script:logFolder = Join-Path $PSScriptRoot "logs"

# ログフォルダ作成
if (!(Test-Path $script:logFolder)) {
    New-Item -ItemType Directory -Path $script:logFolder | Out-Null
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " フォルダナビゲーター v$script:version 起動中..." -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# メインウィンドウXAML
$mainXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Folder Navigator v3.0 - Phase 3 Complete" Width="1200" Height="800"
    WindowStartupLocation="CenterScreen" AllowDrop="True">
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
                <TextBlock Text="Folder Navigator + Batch Rename + Drag&amp;Drop" FontSize="18" FontWeight="Bold" 
                          Foreground="White" Margin="0,0,0,10"/>
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" Text="Folder Path:" Foreground="White" 
                              VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <TextBox Name="txtPath" Grid.Column="1" FontSize="12" Padding="5" AllowDrop="True"/>
                    <Button Name="btnBrowse" Grid.Column="2" Content="Browse" Padding="10,5" 
                           Margin="5" MinWidth="80"/>
                    <Button Name="btnLoad" Grid.Column="3" Content="Load" Padding="10,5" 
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
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="Search:" VerticalAlignment="Center" Margin="0,0,10,0"/>
                <TextBox Name="txtSearch" Grid.Column="1" FontSize="12" Padding="5"/>
                <TextBlock Grid.Column="2" Text="History:" VerticalAlignment="Center" Margin="20,0,10,0"/>
                <ComboBox Name="cmbHistory" Grid.Column="3" MaxWidth="400" HorizontalAlignment="Left"/>
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
                    <TextBlock Grid.Row="0" Text="Folder Tree" FontSize="14" FontWeight="Bold" 
                              Padding="10,5" Background="#2c3e50" Foreground="White"/>
                    <TreeView Name="treeView" Grid.Row="1" FontSize="12" Padding="5"/>
                </Grid>
            </Border>
            
            <!-- スプリッター -->
            <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Stretch" 
                         VerticalAlignment="Stretch" Background="#95a5a6"/>
            
            <!-- 右パネル: ファイル一覧（D&D対応） -->
            <Border Grid.Column="2" BorderBrush="#bdc3c7" BorderThickness="1" Margin="5" AllowDrop="True" Name="dropZone">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <TextBlock Grid.Row="0" Text="Files and Folders (Drop Excel/Text here)" FontSize="14" FontWeight="Bold" 
                              Padding="10,5" Background="#2c3e50" Foreground="White"/>
                    <DataGrid Name="dataGrid" Grid.Row="1" AutoGenerateColumns="False" 
                             CanUserAddRows="False" GridLinesVisibility="None" 
                             AlternatingRowBackground="#f8f9fa" AllowDrop="True">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="80"/>
                            <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="2*"/>
                            <DataGridTextColumn Header="Size" Binding="{Binding Size}" Width="100"/>
                            <DataGridTextColumn Header="Modified" Binding="{Binding Modified}" Width="150"/>
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
                <TextBlock Name="txtStatus" Grid.Column="0" Text="Ready" VerticalAlignment="Center"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <Button Name="btnOpenFolder" Content="Open Folder" Padding="10,5" Margin="5" MinWidth="100"/>
                    <Button Name="btnNewFolder" Content="New Folder" Padding="10,5" Margin="5" MinWidth="100"/>
                    <Button Name="btnTestFolder" Content="Test Folder" Padding="10,5" Margin="5" MinWidth="100"
                           ToolTip="試験項目フォルダを自動作成"/>
                    <Button Name="btnBatchRename" Content="Batch Rename" Padding="10,5" Margin="5" MinWidth="100" 
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
$mainWindow = [Windows.Markup.XamlReader]::Load($reader)

# コントロール取得
$txtPath = $mainWindow.FindName("txtPath")
$btnBrowse = $mainWindow.FindName("btnBrowse")
$btnLoad = $mainWindow.FindName("btnLoad")
$cmbHistory = $mainWindow.FindName("cmbHistory")
$txtSearch = $mainWindow.FindName("txtSearch")
$treeView = $mainWindow.FindName("treeView")
$dataGrid = $mainWindow.FindName("dataGrid")
$dropZone = $mainWindow.FindName("dropZone")
$txtStatus = $mainWindow.FindName("txtStatus")
$btnOpenFolder = $mainWindow.FindName("btnOpenFolder")
$btnNewFolder = $mainWindow.FindName("btnNewFolder")
$btnTestFolder = $mainWindow.FindName("btnTestFolder")
$btnBatchRename = $mainWindow.FindName("btnBatchRename")

# ===============================================
# ログ出力関数
# ===============================================
function Write-RenameLog {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # ログファイル名（タイムスタンプ付き）
    $logFileName = "rename_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".log"
    $logFilePath = Join-Path $script:logFolder $logFileName
    
    # ログファイルに追記
    Add-Content -Path $logFilePath -Value $logMessage -Encoding UTF8
    
    # コンソールにも出力
    switch ($Level) {
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        default { Write-Host $logMessage -ForegroundColor White }
    }
    
    return $logFilePath
}

# ===============================================
# 一括リネーム機能（実運用版）
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
    $form.Text = "一括リネームツール（実運用版）"
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
    $script:previewItems = @()  # プレビューデータを保存
    
    $btnPreview.Add_Click({
        $listBox.Items.Clear()
        $script:previewItems = @()
        
        $items = Get-ChildItem -Path $currentPath -ErrorAction SilentlyContinue
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
    
    # 実行処理（実際のリネーム - Phase 3）
    $btnExecute.Add_Click({
        $result = [System.Windows.Forms.MessageBox]::Show(
            "本当にリネームを実行しますか？`n`n⚠️ この操作は取り消せません`n⚠️ エラーが発生した場合は即座に中止されます",
            "最終確認",
            "YesNo",
            "Warning"
        )
        
        if ($result -eq "Yes") {
            # ログファイルのパスを初期化
            $logFilePath = ""
            
            try {
                # ログ開始
                $logFilePath = Write-RenameLog "========== リネーム処理開始 ==========" "INFO"
                Write-RenameLog "対象フォルダ: $currentPath" "INFO"
                Write-RenameLog "リネーム方式: $($cmbType.SelectedItem)" "INFO"
                Write-RenameLog "処理対象件数: $($script:previewItems.Count)" "INFO"
                Write-RenameLog "" "INFO"
                
                $successCount = 0
                $totalCount = $script:previewItems.Count
                
                # 実際のリネーム処理
                foreach ($item in $script:previewItems) {
                    $oldPath = $item.OldPath
                    $newPath = Join-Path (Split-Path $oldPath -Parent) $item.NewName
                    
                    Write-RenameLog "処理中: $($item.OldName) → $($item.NewName)" "INFO"
                    
                    # 同名ファイルチェック
                    if (Test-Path $newPath) {
                        $errorMsg = "エラー: 同名のファイル/フォルダが既に存在します - $($item.NewName)"
                        Write-RenameLog $errorMsg "ERROR"
                        throw $errorMsg
                    }
                    
                    # リネーム実行
                    try {
                        Rename-Item -Path $oldPath -NewName $item.NewName -ErrorAction Stop
                        $successCount++
                        Write-RenameLog "成功: $($item.OldName) → $($item.NewName)" "SUCCESS"
                    }
                    catch {
                        $errorMsg = "エラー: リネーム失敗 - $($item.OldName)`n詳細: $($_.Exception.Message)"
                        Write-RenameLog $errorMsg "ERROR"
                        throw $errorMsg
                    }
                }
                
                # 全て成功
                Write-RenameLog "" "INFO"
                Write-RenameLog "========== リネーム処理完了 ==========" "SUCCESS"
                Write-RenameLog "成功: $successCount / $totalCount 件" "SUCCESS"
                
                [System.Windows.Forms.MessageBox]::Show(
                    "リネームが完了しました！`n`n成功: $successCount / $totalCount 件`n`nログファイル:`n$logFilePath",
                    "完了",
                    "OK",
                    "Information"
                )
                
                $form.Close()
                
                # ファイルリスト更新
                Load-FileList $currentPath
            }
            catch {
                # エラー発生時
                Write-RenameLog "" "ERROR"
                Write-RenameLog "========== リネーム処理中断 ==========" "ERROR"
                Write-RenameLog "成功: $successCount / $totalCount 件（途中で中止）" "WARNING"
                Write-RenameLog "エラー内容: $_" "ERROR"
                
                [System.Windows.Forms.MessageBox]::Show(
                    "エラーが発生したため処理を中止しました`n`n成功: $successCount / $totalCount 件`n`nエラー内容:`n$_`n`nログファイル:`n$logFilePath",
                    "エラー",
                    "OK",
                    "Error"
                )
                
                # ファイルリスト更新
                Load-FileList $currentPath
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
        $script:currentPath = $path
        
        # 履歴に追加
        if ($script:history -notcontains $path) {
            $script:history = @($path) + $script:history
            if ($script:history.Count -gt $script:maxHistory) {
                $script:history = $script:history[0..($script:maxHistory - 1)]
            }
            Update-HistoryComboBox
        }
        
        $rootItem = New-Object System.Windows.Controls.TreeViewItem
        $rootItem.Header = Split-Path $path -Leaf
        if (!$rootItem.Header) { $rootItem.Header = $path }
        $rootItem.Tag = $path
        
        # ドライブレベルの場合
        if ($path.Length -le 3) {
            Load-SubFolders -parentItem $rootItem -path $path -depth 0 -maxDepth 2
        }
        else {
            Load-SubFolders -parentItem $rootItem -path $path -depth 0 -maxDepth 1
        }
        
        $rootItem.IsExpanded = $true
        $treeView.Items.Add($rootItem)
        
        Load-FileList $path
        $txtStatus.Text = "読み込み完了: $path"
    }
    catch {
        $txtStatus.Text = "エラー: $_"
        Write-Host "エラー: $_" -ForegroundColor Red
    }
}

# サブフォルダ読み込み
function Load-SubFolders {
    param(
        $parentItem,
        $path,
        $depth,
        $maxDepth
    )
    
    if ($depth -ge $maxDepth) {
        $dummyItem = New-Object System.Windows.Controls.TreeViewItem
        $dummyItem.Header = "読み込み中..."
        $parentItem.Items.Add($dummyItem)
        return
    }
    
    try {
        $folders = Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue | 
                   Sort-Object Name
        
        foreach ($folder in $folders) {
            $item = New-Object System.Windows.Controls.TreeViewItem
            $item.Header = $folder.Name
            $item.Tag = $folder.FullName
            
            if ($depth + 1 -lt $maxDepth) {
                Load-SubFolders -parentItem $item -path $folder.FullName -depth ($depth + 1) -maxDepth $maxDepth
            }
            else {
                $hasSubFolders = (Get-ChildItem -Path $folder.FullName -Directory -ErrorAction SilentlyContinue | Measure-Object).Count -gt 0
                if ($hasSubFolders) {
                    $dummyItem = New-Object System.Windows.Controls.TreeViewItem
                    $dummyItem.Header = "読み込み中..."
                    $item.Items.Add($dummyItem)
                }
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
    
    $dataGrid.Items.Clear()
    
    try {
        $folders = Get-ChildItem -Path $path -Directory -ErrorAction SilentlyContinue | Sort-Object Name
        $files = Get-ChildItem -Path $path -File -ErrorAction SilentlyContinue | Sort-Object Name
        
        # フォルダ追加
        foreach ($folder in $folders) {
            $item = New-Object FileItem
            $item.Type = "フォルダ"
            $item.Name = $folder.Name
            $item.Size = "-"
            $item.Modified = $folder.LastWriteTime.ToString("yyyy/MM/dd HH:mm")
            $item.FullPath = $folder.FullName
            $dataGrid.Items.Add($item)
        }
        
        # ファイル追加
        foreach ($file in $files) {
            $item = New-Object FileItem
            $item.Type = "ファイル"
            $item.Name = $file.Name
            
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
# ドラッグ&ドロップ機能（Phase 3 新機能）
# ===============================================

# 対応ファイル形式チェック
function Test-SupportedFileType {
    param($filePath)
    
    $supportedExtensions = @(".xlsx", ".xls", ".xlsm", ".txt", ".csv")
    $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
    
    return $supportedExtensions -contains $extension
}

# DragEnterイベント（視覚フィードバック）
$mainWindow.Add_DragEnter({
    param($sender, $e)
    
    if (!$script:currentPath) {
        $e.Effects = [System.Windows.DragDropEffects]::None
        return
    }
    
    if ($e.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $files = $e.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
        
        # サポートされているファイル形式かチェック
        $hasValidFile = $false
        foreach ($file in $files) {
            if ((Test-Path $file -PathType Leaf) -and (Test-SupportedFileType $file)) {
                $hasValidFile = $true
                break
            }
        }
        
        if ($hasValidFile) {
            $e.Effects = [System.Windows.DragDropEffects]::Copy
        } else {
            $e.Effects = [System.Windows.DragDropEffects]::None
        }
    }
    else {
        $e.Effects = [System.Windows.DragDropEffects]::None
    }
})

# Dropイベント（ファイルコピー）
$mainWindow.Add_Drop({
    param($sender, $e)
    
    if (!$script:currentPath) {
        [System.Windows.MessageBox]::Show(
            "先にフォルダを選択してください", 
            "情報", "OK", "Information")
        return
    }
    
    if ($e.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
        $files = $e.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
        
        $copiedFiles = @()
        $errors = @()
        
        foreach ($file in $files) {
            # ファイルのみ処理（フォルダは無視）
            if (Test-Path $file -PathType Leaf) {
                # サポートされているファイル形式かチェック
                if (Test-SupportedFileType $file) {
                    $fileName = [System.IO.Path]::GetFileName($file)
                    $destination = Join-Path $script:currentPath $fileName
                    
                    try {
                        # 同名ファイルがある場合は確認
                        if (Test-Path $destination) {
                            $result = [System.Windows.MessageBox]::Show(
                                "同名のファイルが既に存在します:`n$fileName`n`n上書きしますか？",
                                "確認",
                                "YesNo",
                                "Question"
                            )
                            
                            if ($result -eq "No") {
                                continue
                            }
                        }
                        
                        # ファイルコピー
                        Copy-Item -Path $file -Destination $destination -Force
                        $copiedFiles += $fileName
                    }
                    catch {
                        $errors += "$fileName : $_"
                    }
                }
            }
        }
        
        # 結果表示
        if ($copiedFiles.Count -gt 0) {
            $message = "ファイルをコピーしました:`n`n" + ($copiedFiles -join "`n")
            
            if ($errors.Count -gt 0) {
                $message += "`n`nエラー:`n" + ($errors -join "`n")
            }
            
            [System.Windows.MessageBox]::Show(
                $message,
                "コピー完了",
                "OK",
                "Information"
            )
            
            # ファイルリスト更新
            Load-FileList $script:currentPath
        }
        elseif ($errors.Count -gt 0) {
            [System.Windows.MessageBox]::Show(
                "エラーが発生しました:`n`n" + ($errors -join "`n"),
                "エラー",
                "OK",
                "Error"
            )
        }
        else {
            [System.Windows.MessageBox]::Show(
                "対応していないファイル形式です`n`n対応形式: Excel (.xlsx, .xls, .xlsm), Text (.txt, .csv)",
                "情報",
                "OK",
                "Information"
            )
        }
    }
})

# DataGridにもD&D機能を追加
$dataGrid.Add_DragEnter({
    param($sender, $e)
    $mainWindow_DragEnter.Invoke($sender, $e)
})

$dataGrid.Add_Drop({
    param($sender, $e)
    $mainWindow_Drop.Invoke($sender, $e)
})

# DropZoneにもD&D機能を追加
$dropZone.Add_DragEnter({
    param($sender, $e)
    $mainWindow_DragEnter.Invoke($sender, $e)
})

$dropZone.Add_Drop({
    param($sender, $e)
    $mainWindow_Drop.Invoke($sender, $e)
})

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

# 検索機能
$txtSearch.Add_TextChanged({
    # TODO: 検索機能の実装
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
Write-Host "  ✅ 一括リネーム機能（実運用版）" -ForegroundColor White
Write-Host "  ✅ ドラッグ&ドロップ（Excel/Text）" -ForegroundColor White
Write-Host "  ✅ 履歴機能" -ForegroundColor White
Write-Host ""
Write-Host "【ログフォルダ】" -ForegroundColor Yellow
Write-Host "  $script:logFolder" -ForegroundColor White
Write-Host ""

# メインウィンドウ表示
$mainWindow.ShowDialog() | Out-Null
Write-Host "アプリケーションを終了しました" -ForegroundColor Yellow
