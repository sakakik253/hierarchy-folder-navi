# ===============================================
# フォルダナビゲーター Phase 3.6 Complete
# Version: 3.6.0
# Date: 2024-11-23
# Author: KENJI
# ===============================================
# 
# 【機能一覧】
# - フォルダツリー表示（階層構造）
# - ファイル・フォルダ一覧表示
# - リアルタイム検索（再帰検索）
# - ブックマーク機能
#   - よく使うフォルダを登録
#   - ワンクリックで移動
#   - 追加・削除可能
#   - 自動保存
# - Excel/CSV出力機能
#   - フォルダ構造をExcel/CSV出力
#   - サブフォルダ再帰対応
#   - ファイルサイズ・更新日時出力
#   - 出力後に自動で開く
# - ファイル削除機能
#   - 選択したファイル/フォルダを削除
#   - 複数選択対応
#   - 削除前に確認ダイアログ
#   - 右クリックメニュー対応
#   - Deleteキー対応
# - ファイル移動/コピー機能（新機能！）
#   - コピー/切り取り/貼り付け
#   - フォルダに移動
#   - キーボードショートカット対応（Ctrl+C/X/V）
#   - 右クリックメニュー対応
#   - 同名ファイル上書き確認
# - 新規フォルダ作成（試験項目連番対応）
# - フォルダを開く（エクスプローラー起動）
# - 一括リネーム機能
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
# - UI完全日本語化
# - スクロール機能
# ===============================================

# STAモードチェック（WPF必須）
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Host "========================================" -ForegroundColor Yellow
    Write-Host " WPFにはSTAモードが必要です" -ForegroundColor Yellow
    Write-Host " STAモードで再起動します..." -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    
    # STAモードで再起動
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
$script:version = "3.6.0"
$script:previewItems = @()
$script:bookmarks = @()
$script:bookmarkFile = Join-Path $PSScriptRoot "bookmarks.json"
# クリップボード管理用
$script:clipboard = @{
    Items = @()
    Operation = ""  # "Copy" or "Cut"
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " フォルダナビゲーター v$script:version 起動中..." -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# メインウィンドウXAML
$mainXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="フォルダナビゲーター v3.6 - Phase 3.6" Width="1400" Height="800"
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
                <TextBlock Text="フォルダナビゲーター + リネーム + コピー/移動 + Excel + 削除" FontSize="18" FontWeight="Bold" 
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
                <ColumnDefinition Width="200" MinWidth="150"/>
                <ColumnDefinition Width="5"/>
                <ColumnDefinition Width="2*" MinWidth="300"/>
                <ColumnDefinition Width="5"/>
                <ColumnDefinition Width="3*" MinWidth="400"/>
            </Grid.ColumnDefinitions>
            
            <!-- 左パネル: ブックマーク -->
            <Border Grid.Column="0" BorderBrush="#bdc3c7" BorderThickness="1" Margin="5">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <TextBlock Grid.Row="0" Text="★ ブックマーク" FontSize="14" FontWeight="Bold" 
                              Padding="10,5" Background="#e67e22" Foreground="White"/>
                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                        <ListBox Name="lstBookmarks" FontSize="11" BorderThickness="0"
                                 SelectionMode="Single" Padding="5">
                            <ListBox.ItemTemplate>
                                <DataTemplate>
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="Auto"/>
                                        </Grid.ColumnDefinitions>
                                        <TextBlock Grid.Column="0" Text="{Binding Name}" 
                                                  TextTrimming="CharacterEllipsis"
                                                  ToolTip="{Binding Path}"/>
                                        <Button Grid.Column="1" Content="×" Width="20" Height="20" 
                                               FontSize="10" Margin="5,0,0,0"
                                               Tag="{Binding Path}"
                                               ToolTip="削除"/>
                                    </Grid>
                                </DataTemplate>
                            </ListBox.ItemTemplate>
                        </ListBox>
                    </ScrollViewer>
                    <Button Name="btnAddBookmark" Grid.Row="2" Content="+ 現在のフォルダを追加" 
                           Padding="5" Margin="5" FontSize="10" Background="#3498db" 
                           Foreground="White" FontWeight="Bold"/>
                </Grid>
            </Border>
            
            <!-- スプリッター1 -->
            <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Stretch" 
                         VerticalAlignment="Stretch" Background="#95a5a6"/>
            
            <!-- 中央パネル: フォルダツリー -->
            <Border Grid.Column="2" BorderBrush="#bdc3c7" BorderThickness="1" Margin="5">
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
            
            <!-- スプリッター2 -->
            <GridSplitter Grid.Column="3" Width="5" HorizontalAlignment="Stretch" 
                         VerticalAlignment="Stretch" Background="#95a5a6"/>
            
            <!-- 右パネル: ファイル一覧 -->
            <Border Grid.Column="4" BorderBrush="#bdc3c7" BorderThickness="1" Margin="5">
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
                             SelectionMode="Extended"
                             VerticalScrollBarVisibility="Auto"
                             HorizontalScrollBarVisibility="Auto"
                             EnableRowVirtualization="True"
                             EnableColumnVirtualization="True">
                        <DataGrid.ContextMenu>
                            <ContextMenu>
                                <MenuItem Name="menuCopy" Header="コピー (Ctrl+C)" />
                                <MenuItem Name="menuCut" Header="切り取り (Ctrl+X)" />
                                <MenuItem Name="menuPaste" Header="貼り付け (Ctrl+V)" />
                                <Separator/>
                                <MenuItem Name="menuDelete" Header="削除 (Delete)" />
                            </ContextMenu>
                        </DataGrid.ContextMenu>
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
                    <Button Name="btnCopy" Content="コピー" Padding="10,5" Margin="5" MinWidth="80"
                           ToolTip="選択したファイル/フォルダをコピー"/>
                    <Button Name="btnCut" Content="切り取り" Padding="10,5" Margin="5" MinWidth="80"
                           ToolTip="選択したファイル/フォルダを切り取り"/>
                    <Button Name="btnPaste" Content="貼り付け" Padding="10,5" Margin="5" MinWidth="80"
                           ToolTip="コピー/切り取りしたファイルを貼り付け"/>
                    <Button Name="btnMoveToFolder" Content="フォルダに移動..." Padding="10,5" Margin="5" MinWidth="120"
                           Background="#3498db" Foreground="White"
                           ToolTip="選択したファイル/フォルダを別フォルダに移動"/>
                    <Button Name="btnExportExcel" Content="Excel出力" Padding="10,5" Margin="5" MinWidth="100"
                           Background="#27ae60" Foreground="White" FontWeight="Bold"
                           ToolTip="現在のフォルダ構造をExcel/CSVに出力"/>
                    <Button Name="btnDelete" Content="削除" Padding="10,5" Margin="5" MinWidth="80"
                           Background="#e67e22" Foreground="White" FontWeight="Bold"
                           ToolTip="選択したファイル/フォルダを削除"/>
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

# Bookmarkクラス
Add-Type @"
public class BookmarkItem {
    public string Name { get; set; }
    public string Path { get; set; }
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
$btnCopy = $mainWindow.FindName("btnCopy")
$btnCut = $mainWindow.FindName("btnCut")
$btnPaste = $mainWindow.FindName("btnPaste")
$btnMoveToFolder = $mainWindow.FindName("btnMoveToFolder")
$btnExportExcel = $mainWindow.FindName("btnExportExcel")
$btnDelete = $mainWindow.FindName("btnDelete")
$btnBatchRename = $mainWindow.FindName("btnBatchRename")
$lstBookmarks = $mainWindow.FindName("lstBookmarks")
$btnAddBookmark = $mainWindow.FindName("btnAddBookmark")
$menuCopy = $mainWindow.FindName("menuCopy")
$menuCut = $mainWindow.FindName("menuCut")
$menuPaste = $mainWindow.FindName("menuPaste")
$menuDelete = $mainWindow.FindName("menuDelete")

# ===============================================
# ブックマーク機能
# ===============================================

# ブックマーク読み込み
function Load-Bookmarks {
    if (Test-Path $script:bookmarkFile) {
        try {
            $json = Get-Content $script:bookmarkFile -Raw -Encoding UTF8 | ConvertFrom-Json
            $script:bookmarks = @()
            foreach ($item in $json) {
                $bookmark = New-Object BookmarkItem
                $bookmark.Name = $item.Name
                $bookmark.Path = $item.Path
                $script:bookmarks += $bookmark
            }
            Update-BookmarkList
            Write-Host "ブックマーク読み込み完了: $($script:bookmarks.Count)件" -ForegroundColor Green
        }
        catch {
            Write-Host "ブックマーク読み込みエラー: $_" -ForegroundColor Red
            $script:bookmarks = @()
        }
    }
    else {
        Write-Host "ブックマークファイルが見つかりません。新規作成します。" -ForegroundColor Yellow
        $script:bookmarks = @()
    }
}

# ブックマーク保存
function Save-Bookmarks {
    try {
        $json = $script:bookmarks | ConvertTo-Json -Depth 10
        [System.IO.File]::WriteAllText($script:bookmarkFile, $json, [System.Text.UTF8Encoding]::new($false))
        Write-Host "ブックマーク保存完了: $($script:bookmarks.Count)件" -ForegroundColor Green
    }
    catch {
        Write-Host "ブックマーク保存エラー: $_" -ForegroundColor Red
    }
}

# ブックマークリスト更新
function Update-BookmarkList {
    $lstBookmarks.Items.Clear()
    foreach ($bookmark in $script:bookmarks) {
        $lstBookmarks.Items.Add($bookmark)
    }
}

# ブックマーク追加
function Add-Bookmark {
    param($path)
    
    if (!$path -or !(Test-Path $path)) {
        [System.Windows.MessageBox]::Show(
            "有効なフォルダを選択してください",
            "エラー",
            "OK",
            "Warning"
        )
        return
    }
    
    # 既に登録されているかチェック
    $exists = $script:bookmarks | Where-Object { $_.Path -eq $path }
    if ($exists) {
        [System.Windows.MessageBox]::Show(
            "このフォルダは既にブックマークに登録されています",
            "情報",
            "OK",
            "Information"
        )
        return
    }
    
    # 名前を入力
    $folderName = Split-Path $path -Leaf
    if (!$folderName) { $folderName = $path }
    
    $name = [Microsoft.VisualBasic.Interaction]::InputBox(
        "ブックマーク名を入力してください:",
        "ブックマーク追加",
        $folderName
    )
    
    if ($name -eq "") { return }
    
    # ブックマーク追加
    $bookmark = New-Object BookmarkItem
    $bookmark.Name = $name
    $bookmark.Path = $path
    
    $script:bookmarks += $bookmark
    Save-Bookmarks
    Update-BookmarkList
    
    [System.Windows.MessageBox]::Show(
        "ブックマークを追加しました",
        "成功",
        "OK",
        "Information"
    )
}

# ブックマーク削除
function Remove-Bookmark {
    param($path)
    
    $result = [System.Windows.MessageBox]::Show(
        "このブックマークを削除しますか？",
        "確認",
        "YesNo",
        "Question"
    )
    
    if ($result -eq "Yes") {
        $script:bookmarks = $script:bookmarks | Where-Object { $_.Path -ne $path }
        Save-Bookmarks
        Update-BookmarkList
    }
}

# ブックマーククリックイベント
$lstBookmarks.Add_SelectionChanged({
    if ($lstBookmarks.SelectedItem) {
        $bookmark = $lstBookmarks.SelectedItem
        if (Test-Path $bookmark.Path) {
            $txtPath.Text = $bookmark.Path
            Load-FolderTree $bookmark.Path
        }
        else {
            [System.Windows.MessageBox]::Show(
                "フォルダが見つかりません:`n$($bookmark.Path)",
                "エラー",
                "OK",
                "Error"
            )
        }
    }
})

# ブックマーク追加ボタン
$btnAddBookmark.Add_Click({
    Add-Bookmark $script:currentPath
})

# ブックマーク削除ボタン（ListBox内のボタン）
$lstBookmarks.AddHandler(
    [System.Windows.Controls.Primitives.ButtonBase]::ClickEvent,
    [System.Windows.RoutedEventHandler]{
        param($sender, $e)
        if ($e.OriginalSource.GetType().Name -eq "Button") {
            $path = $e.OriginalSource.Tag
            if ($path) {
                Remove-Bookmark $path
                $e.Handled = $true
            }
        }
    }
)

# ===============================================
# ドラッグ&ドロップ機能
# ===============================================

# サポートする拡張子
$script:supportedExtensions = @('.xlsx', '.xls', '.xlsm', '.txt', '.csv')

# ドラッグエンター
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

# ドロップ処理
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
            
            if ($script:supportedExtensions -contains $ext) {
                $fileName = [System.IO.Path]::GetFileName($file)
                $destPath = Join-Path $script:currentPath $fileName
                
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
        
        if ($copiedCount -gt 0) {
            $message = "コピー完了: $copiedCount 件"
            if ($skippedCount -gt 0) {
                $message += " (スキップ: $skippedCount 件)"
            }
            $txtStatus.Text = $message
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
    $mainWindow.RaiseEvent($e)
})

# ===============================================
# 一括リネーム機能（Phase 3.2と同じ）
# ===============================================
function Show-BatchRenameDialog {
    param($currentPath)
    
    if (!$currentPath) {
        [System.Windows.Forms.MessageBox]::Show(
            "フォルダを選択してください",
            "エラー", "OK", "Warning")
        return
    }
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "一括リネームツール"
    $form.Size = New-Object System.Drawing.Size(750, 650)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("メイリオ", 10)
    
    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Location = New-Object System.Drawing.Point(20, 10)
    $lblInfo.Size = New-Object System.Drawing.Size(700, 60)
    $lblInfo.Text = "ファイル・フォルダの名前を一括で変更します`n現在のフォルダ: $currentPath"
    $lblInfo.BackColor = [System.Drawing.Color]::LightBlue
    $lblInfo.Padding = New-Object System.Windows.Forms.Padding(10)
    
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
    
    $grpUsage = New-Object System.Windows.Forms.GroupBox
    $grpUsage.Location = New-Object System.Drawing.Point(470, 85)
    $grpUsage.Size = New-Object System.Drawing.Size(250, 120)
    $grpUsage.Text = "【使い方】"
    
    $lblUsage = New-Object System.Windows.Forms.Label
    $lblUsage.Location = New-Object System.Drawing.Point(10, 20)
    $lblUsage.Size = New-Object System.Drawing.Size(230, 90)
    $lblUsage.Text = "1. リネーム方式を選択`n2. 文字を入力`n3. プレビューで確認`n4. 実行ボタンでリネーム"
    
    $grpUsage.Controls.Add($lblUsage)
    
    $lblInput1 = New-Object System.Windows.Forms.Label
    $lblInput1.Location = New-Object System.Drawing.Point(20, 130)
    $lblInput1.Size = New-Object System.Drawing.Size(120, 25)
    $lblInput1.Text = "追加する文字:"
    
    $txtInput1 = New-Object System.Windows.Forms.TextBox
    $txtInput1.Location = New-Object System.Drawing.Point(150, 127)
    $txtInput1.Size = New-Object System.Drawing.Size(300, 30)
    $txtInput1.Text = "2024_"
    
    $lblInput2 = New-Object System.Windows.Forms.Label
    $lblInput2.Location = New-Object System.Drawing.Point(20, 165)
    $lblInput2.Size = New-Object System.Drawing.Size(120, 25)
    $lblInput2.Text = "置換後の文字:"
    $lblInput2.Visible = $false
    
    $txtInput2 = New-Object System.Windows.Forms.TextBox
    $txtInput2.Location = New-Object System.Drawing.Point(150, 162)
    $txtInput2.Size = New-Object System.Drawing.Size(300, 30)
    $txtInput2.Visible = $false
    
    $btnPreview = New-Object System.Windows.Forms.Button
    $btnPreview.Location = New-Object System.Drawing.Point(150, 210)
    $btnPreview.Size = New-Object System.Drawing.Size(150, 40)
    $btnPreview.Text = "プレビュー"
    $btnPreview.BackColor = [System.Drawing.Color]::LightGreen
    $btnPreview.Font = New-Object System.Drawing.Font("メイリオ", 11, [System.Drawing.FontStyle]::Bold)
    
    $lblPreview = New-Object System.Windows.Forms.Label
    $lblPreview.Location = New-Object System.Drawing.Point(20, 265)
    $lblPreview.Size = New-Object System.Drawing.Size(700, 25)
    $lblPreview.Text = "変更プレビュー（変更前 → 変更後）:"
    $lblPreview.Font = New-Object System.Drawing.Font("メイリオ", 10, [System.Drawing.FontStyle]::Bold)
    
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(20, 295)
    $listBox.Size = New-Object System.Drawing.Size(700, 250)
    $listBox.Font = New-Object System.Drawing.Font("MS Gothic", 10)
    $listBox.HorizontalScrollbar = $true
    
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
    
    $cmbType.Add_SelectedIndexChanged({
        switch ($cmbType.SelectedIndex) {
            0 {
                $lblInput1.Text = "追加する文字:"
                $txtInput1.Text = "2024_"
                $lblInput2.Visible = $false
                $txtInput2.Visible = $false
                $lblUsage.Text = "ファイル名の先頭に`n文字を追加します`n`n例: test.txt`n→ 2024_test.txt"
            }
            1 {
                $lblInput1.Text = "追加する文字:"
                $txtInput1.Text = "_完了"
                $lblInput2.Visible = $false
                $txtInput2.Visible = $false
                $lblUsage.Text = "ファイル名の末尾に`n文字を追加します`n`n例: test.txt`n→ test_完了.txt"
            }
            2 {
                $lblInput1.Text = "置換前の文字:"
                $txtInput1.Text = "2025"
                $lblInput2.Visible = $true
                $txtInput2.Visible = $true
                $txtInput2.Text = "2024"
                $lblUsage.Text = "文字列を置換します`n`n例: 2025_test.txt`n→ 2024_test.txt"
            }
            3 {
                $lblInput1.Text = "プレフィックス:"
                $txtInput1.Text = "File_"
                $lblInput2.Visible = $false
                $txtInput2.Visible = $false
                $lblUsage.Text = "連番を追加します`n`n例: test.txt`n→ File_001_test.txt"
            }
        }
    })
    
    $btnPreview.Add_Click({
        $listBox.Items.Clear()
        $script:previewItems = @()
        $items = Get-ChildItem -Path $currentPath -ErrorAction SilentlyContinue | Select-Object -First 30
        $counter = 1
        
        foreach ($item in $items) {
            $newName = ""
            $type = if ($item.PSIsContainer) { "[フォルダ]" } else { "[ファイル]" }
            
            switch ($cmbType.SelectedIndex) {
                0 {
                    $newName = $txtInput1.Text + $item.Name
                }
                1 {
                    if ($item.PSIsContainer) {
                        $newName = $item.Name + $txtInput1.Text
                    } else {
                        $name = [System.IO.Path]::GetFileNameWithoutExtension($item.Name)
                        $ext = [System.IO.Path]::GetExtension($item.Name)
                        $newName = $name + $txtInput1.Text + $ext
                    }
                }
                2 {
                    if ($txtInput2.Visible) {
                        $newName = $item.Name.Replace($txtInput1.Text, $txtInput2.Text)
                    } else {
                        $newName = $item.Name
                    }
                }
                3 {
                    $newName = $txtInput1.Text + "{0:D3}_" -f $counter + $item.Name
                    $counter++
                }
            }
            
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
                    
                    if ($oldPath -ne $newPath) {
                        Rename-Item -Path $oldPath -NewName $newName -ErrorAction Stop
                        $successCount++
                    }
                }
                
                if ($failCount -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show(
                        "リネームが完了しました！`n`n成功: $successCount 件",
                        "完了",
                        "OK",
                        "Information"
                    )
                    
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
    
    $form.ShowDialog() | Out-Null
}

# ===============================================
# メイン機能関数（Phase 3.2と同じ）
# ===============================================

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

function Load-FileList {
    param($path)
    
    if (!(Test-Path $path)) { return }
    
    $dataGrid.Items.Clear()
    
    try {
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
        
        $files = Get-ChildItem -Path $path -File -ErrorAction SilentlyContinue
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

function Update-HistoryComboBox {
    $cmbHistory.Items.Clear()
    foreach ($item in $script:history) {
        $cmbHistory.Items.Add($item)
    }
}

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

function Create-TestFolder {
    if (!$script:currentPath) {
        [System.Windows.MessageBox]::Show(
            "先にフォルダを選択してください",
            "情報", "OK", "Information")
        return
    }
    
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
# ファイル削除機能
# ===============================================
function Delete-SelectedItems {
    $selectedItems = $dataGrid.SelectedItems
    
    if ($selectedItems.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "削除するファイル/フォルダを選択してください",
            "情報", "OK", "Information")
        return
    }
    
    # 確認ダイアログ
    $itemList = $selectedItems | ForEach-Object { $_.Name }
    $message = "以下の $($selectedItems.Count) 件を削除します:`n`n" + 
               ($itemList | Select-Object -First 10 | ForEach-Object { "  • $_" } | Out-String) +
               $(if ($selectedItems.Count -gt 10) { "`n  ...他 $($selectedItems.Count - 10) 件" } else { "" }) +
               "`n`n⚠️ この操作は取り消せません。本当に削除しますか？"
    
    $result = [System.Windows.MessageBox]::Show(
        $message,
        "削除の確認",
        "YesNo",
        "Warning")
    
    if ($result -eq "Yes") {
        $successCount = 0
        $errorCount = 0
        $errors = @()
        
        foreach ($item in $selectedItems) {
            try {
                $path = $item.FullPath
                
                if (Test-Path $path -PathType Container) {
                    # フォルダの削除
                    Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                } else {
                    # ファイルの削除
                    Remove-Item -Path $path -Force -ErrorAction Stop
                }
                
                $successCount++
                Write-Host "削除完了: $($item.Name)" -ForegroundColor Green
            }
            catch {
                $errorCount++
                $errors += "$($item.Name): $_"
                Write-Host "削除失敗: $($item.Name) - $_" -ForegroundColor Red
            }
        }
        
        # 結果表示
        $resultMsg = "削除完了！`n`n成功: $successCount 件"
        if ($errorCount -gt 0) {
            $resultMsg += "`n失敗: $errorCount 件`n`nエラー:`n" + 
                         ($errors | Select-Object -First 5 | ForEach-Object { "  • $_" } | Out-String) +
                         $(if ($errors.Count -gt 5) { "`n  ...他 $($errors.Count - 5) 件" } else { "" })
        }
        
        $msgTitle = if ($errorCount -eq 0) { "完了" } else { "一部エラー" }
        $msgIcon = if ($errorCount -eq 0) { "Information" } else { "Warning" }
        
        [System.Windows.MessageBox]::Show(
            $resultMsg,
            $msgTitle,
            "OK",
            $msgIcon)
        
        # リスト更新
        Load-FileList $script:currentPath
    }
}

# ===============================================
# ファイル移動/コピー機能
# ===============================================

# コピー/切り取り機能
function Copy-SelectedItems {
    param([bool]$isCut = $false)
    
    $selectedItems = $dataGrid.SelectedItems
    
    if ($selectedItems.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "コピーするファイル/フォルダを選択してください",
            "情報", "OK", "Information")
        return
    }
    
    $script:clipboard.Items = @($selectedItems | ForEach-Object { $_.FullPath })
    $script:clipboard.Operation = if ($isCut) { "Cut" } else { "Copy" }
    
    $operation = if ($isCut) { "切り取り" } else { "コピー" }
    $txtStatus.Text = "$operation : $($script:clipboard.Items.Count) 件"
    Write-Host "$operation : $($script:clipboard.Items.Count) 件" -ForegroundColor Cyan
}

# 貼り付け機能
function Paste-Items {
    if ($script:clipboard.Items.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "コピー/切り取りされたファイルがありません",
            "情報", "OK", "Information")
        return
    }
    
    if (!$script:currentPath) {
        [System.Windows.MessageBox]::Show(
            "貼り付け先のフォルダを選択してください",
            "情報", "OK", "Information")
        return
    }
    
    $successCount = 0
    $errorCount = 0
    $errors = @()
    
    foreach ($sourcePath in $script:clipboard.Items) {
        try {
            if (!(Test-Path $sourcePath)) {
                $errors += "$(Split-Path $sourcePath -Leaf): ファイルが見つかりません"
                $errorCount++
                continue
            }
            
            $fileName = Split-Path $sourcePath -Leaf
            $destPath = Join-Path $script:currentPath $fileName
            
            # 同名ファイルチェック
            if (Test-Path $destPath) {
                $result = [System.Windows.MessageBox]::Show(
                    "同名のファイル/フォルダが既に存在します:`n$fileName`n`n上書きしますか？",
                    "確認",
                    "YesNo",
                    "Question")
                
                if ($result -eq "No") {
                    continue
                }
            }
            
            if ($script:clipboard.Operation -eq "Cut") {
                # 移動
                Move-Item -Path $sourcePath -Destination $destPath -Force -ErrorAction Stop
                Write-Host "移動完了: $fileName" -ForegroundColor Green
            } else {
                # コピー
                if (Test-Path $sourcePath -PathType Container) {
                    Copy-Item -Path $sourcePath -Destination $destPath -Recurse -Force -ErrorAction Stop
                } else {
                    Copy-Item -Path $sourcePath -Destination $destPath -Force -ErrorAction Stop
                }
                Write-Host "コピー完了: $fileName" -ForegroundColor Green
            }
            
            $successCount++
        }
        catch {
            $errorCount++
            $errors += "$(Split-Path $sourcePath -Leaf): $_"
            Write-Host "エラー: $(Split-Path $sourcePath -Leaf) - $_" -ForegroundColor Red
        }
    }
    
    # 切り取りの場合はクリップボードをクリア
    if ($script:clipboard.Operation -eq "Cut") {
        $script:clipboard.Items = @()
        $script:clipboard.Operation = ""
    }
    
    # 結果表示
    $operation = if ($script:clipboard.Operation -eq "Cut") { "移動" } else { "コピー" }
    $resultMsg = "$operation 完了！`n`n成功: $successCount 件"
    if ($errorCount -gt 0) {
        $resultMsg += "`n失敗: $errorCount 件`n`nエラー:`n" + 
                     ($errors | Select-Object -First 5 | ForEach-Object { "  • $_" } | Out-String) +
                     $(if ($errors.Count -gt 5) { "`n  ...他 $($errors.Count - 5) 件" } else { "" })
    }
    
    $msgTitle = if ($errorCount -eq 0) { "完了" } else { "一部エラー" }
    $msgIcon = if ($errorCount -eq 0) { "Information" } else { "Warning" }
    
    [System.Windows.MessageBox]::Show(
        $resultMsg,
        $msgTitle,
        "OK",
        $msgIcon)
    
    # リスト更新
    Load-FileList $script:currentPath
    $txtStatus.Text = "準備完了"
}

# フォルダに移動機能
function Move-ToSelectedFolder {
    $selectedItems = $dataGrid.SelectedItems
    
    if ($selectedItems.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "移動するファイル/フォルダを選択してください",
            "情報", "OK", "Information")
        return
    }
    
    # フォルダ選択ダイアログ
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "移動先のフォルダを選択してください"
    $dialog.ShowNewFolderButton = $true
    
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $destFolder = $dialog.SelectedPath
        
        $successCount = 0
        $errorCount = 0
        $errors = @()
        
        foreach ($item in $selectedItems) {
            try {
                $sourcePath = $item.FullPath
                $fileName = $item.Name
                $destPath = Join-Path $destFolder $fileName
                
                # 同名チェック
                if (Test-Path $destPath) {
                    $result = [System.Windows.MessageBox]::Show(
                        "同名のファイル/フォルダが既に存在します:`n$fileName`n`n上書きしますか？",
                        "確認",
                        "YesNo",
                        "Question")
                    
                    if ($result -eq "No") {
                        continue
                    }
                }
                
                Move-Item -Path $sourcePath -Destination $destPath -Force -ErrorAction Stop
                $successCount++
                Write-Host "移動完了: $fileName → $destFolder" -ForegroundColor Green
            }
            catch {
                $errorCount++
                $errors += "$($item.Name): $_"
                Write-Host "移動失敗: $($item.Name) - $_" -ForegroundColor Red
            }
        }
        
        # 結果表示
        $resultMsg = "移動完了！`n`n成功: $successCount 件"
        if ($errorCount -gt 0) {
            $resultMsg += "`n失敗: $errorCount 件`n`nエラー:`n" + 
                         ($errors | Select-Object -First 5 | ForEach-Object { "  • $_" } | Out-String) +
                         $(if ($errors.Count -gt 5) { "`n  ...他 $($errors.Count - 5) 件" } else { "" })
        }
        
        $msgTitle = if ($errorCount -eq 0) { "完了" } else { "一部エラー" }
        $msgIcon = if ($errorCount -eq 0) { "Information" } else { "Warning" }
        
        [System.Windows.MessageBox]::Show(
            $resultMsg,
            $msgTitle,
            "OK",
            $msgIcon)
        
        # リスト更新
        Load-FileList $script:currentPath
    }
}

# ===============================================
# Excel/CSV出力機能
# ===============================================
function Export-FolderStructure {
    param($rootPath)
    
    if (!$rootPath -or !(Test-Path $rootPath)) {
        [System.Windows.MessageBox]::Show(
            "フォルダを選択してください",
            "情報", "OK", "Information")
        return
    }
    
    # 出力設定ダイアログ
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Excel出力設定"
    $form.Size = New-Object System.Drawing.Size(400, 300)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("メイリオ", 10)
    
    # 説明
    $lblInfo = New-Object System.Windows.Forms.Label
    $lblInfo.Location = New-Object System.Drawing.Point(20, 20)
    $lblInfo.Size = New-Object System.Drawing.Size(350, 30)
    $lblInfo.Text = "フォルダ構造をExcel/CSVファイルに出力します"
    
    # オプション
    $chkSubfolders = New-Object System.Windows.Forms.CheckBox
    $chkSubfolders.Location = New-Object System.Drawing.Point(20, 60)
    $chkSubfolders.Size = New-Object System.Drawing.Size(350, 25)
    $chkSubfolders.Text = "サブフォルダを含める"
    $chkSubfolders.Checked = $true
    
    $chkSize = New-Object System.Windows.Forms.CheckBox
    $chkSize.Location = New-Object System.Drawing.Point(20, 90)
    $chkSize.Size = New-Object System.Drawing.Size(350, 25)
    $chkSize.Text = "ファイルサイズを含める"
    $chkSize.Checked = $true
    
    $chkDate = New-Object System.Windows.Forms.CheckBox
    $chkDate.Location = New-Object System.Drawing.Point(20, 120)
    $chkDate.Size = New-Object System.Drawing.Size(350, 25)
    $chkDate.Text = "更新日時を含める"
    $chkDate.Checked = $true
    
    # 出力形式
    $lblFormat = New-Object System.Windows.Forms.Label
    $lblFormat.Location = New-Object System.Drawing.Point(20, 160)
    $lblFormat.Size = New-Object System.Drawing.Size(100, 25)
    $lblFormat.Text = "出力形式："
    
    $cmbFormat = New-Object System.Windows.Forms.ComboBox
    $cmbFormat.Location = New-Object System.Drawing.Point(120, 157)
    $cmbFormat.Size = New-Object System.Drawing.Size(150, 25)
    $cmbFormat.DropDownStyle = "DropDownList"
    $cmbFormat.Items.AddRange(@("CSV", "TSV"))
    $cmbFormat.SelectedIndex = 0
    
    # ボタン
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Location = New-Object System.Drawing.Point(200, 210)
    $btnOK.Size = New-Object System.Drawing.Size(80, 35)
    $btnOK.Text = "出力"
    $btnOK.DialogResult = "OK"
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(290, 210)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 35)
    $btnCancel.Text = "キャンセル"
    $btnCancel.DialogResult = "Cancel"
    
    $form.Controls.AddRange(@($lblInfo, $chkSubfolders, $chkSize, $chkDate, 
                              $lblFormat, $cmbFormat, $btnOK, $btnCancel))
    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnCancel
    
    if ($form.ShowDialog() -eq "OK") {
        # ファイル保存ダイアログ
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = if ($cmbFormat.SelectedItem -eq "CSV") {
            "CSVファイル (*.csv)|*.csv"
        } else {
            "TSVファイル (*.txt)|*.txt"
        }
        $saveDialog.FileName = "フォルダ構造_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        
        if ($saveDialog.ShowDialog() -eq "OK") {
            try {
                $delimiter = if ($cmbFormat.SelectedItem -eq "CSV") { "," } else { "`t" }
                
                # データ収集
                $items = if ($chkSubfolders.Checked) {
                    Get-ChildItem -Path $rootPath -Recurse -ErrorAction SilentlyContinue
                } else {
                    Get-ChildItem -Path $rootPath -ErrorAction SilentlyContinue
                }
                
                # 階層構造でソート
                $items = $items | Sort-Object FullName
                
                # CSV/TSV作成
                $data = @()
                # ヘッダー行（階層列を追加）
                $data += "階層${delimiter}種類${delimiter}名前${delimiter}パス" + 
                         $(if ($chkSize.Checked) { "${delimiter}サイズ" } else { "" }) +
                         $(if ($chkDate.Checked) { "${delimiter}更新日時" } else { "" })
                
                foreach ($item in $items) {
                    $relativePath = $item.FullName.Replace($rootPath, "").TrimStart('\')
                    $type = if ($item.PSIsContainer) { "フォルダ" } else { "ファイル" }
                    
                    # 階層レベルを計算（\の数 = 階層）
                    $level = 0
                    if ($relativePath) {
                        $level = ($relativePath.Split('\').Length) - 1
                    }
                    
                    # 階層に応じてインデントを追加（全角スペース2つ）
                    $indent = "　" * $level
                    $displayName = $indent + $item.Name
                    
                    # 行を作成
                    $line = "$level${delimiter}$type${delimiter}$displayName${delimiter}$relativePath"
                    
                    if ($chkSize.Checked) {
                        if ($item.PSIsContainer) {
                            $line += "${delimiter}-"
                        } else {
                            $sizeStr = if ($item.Length -lt 1KB) { "$($item.Length) B" }
                                      elseif ($item.Length -lt 1MB) { "{0:N1} KB" -f ($item.Length / 1KB) }
                                      elseif ($item.Length -lt 1GB) { "{0:N1} MB" -f ($item.Length / 1MB) }
                                      else { "{0:N1} GB" -f ($item.Length / 1GB) }
                            $line += "${delimiter}$sizeStr"
                        }
                    }
                    
                    if ($chkDate.Checked) {
                        $line += "${delimiter}$($item.LastWriteTime.ToString('yyyy/MM/dd HH:mm'))"
                    }
                    
                    $data += $line
                }
                
                # ファイル保存（UTF-8 BOM）
                $data | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
                
                [System.Windows.MessageBox]::Show(
                    "Excel出力が完了しました！`n`n出力件数: $($items.Count) 件`n保存先: $($saveDialog.FileName)",
                    "完了", "OK", "Information")
                
                # ファイルを開く
                Start-Process $saveDialog.FileName
            }
            catch {
                [System.Windows.MessageBox]::Show(
                    "Excel出力に失敗しました:`n$_",
                    "エラー", "OK", "Error")
            }
        }
    }
}

# ===============================================
# イベントハンドラ
# ===============================================

$btnBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "フォルダを選択してください"
    $dialog.ShowNewFolderButton = $true
    
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtPath.Text = $dialog.SelectedPath
        Load-FolderTree $dialog.SelectedPath
    }
})

$btnLoad.Add_Click({
    $path = $txtPath.Text.Trim()
    if ($path) {
        Load-FolderTree $path
    }
})

$txtPath.Add_KeyDown({
    if ($_.Key -eq "Return") {
        $path = $txtPath.Text.Trim()
        if ($path) {
            Load-FolderTree $path
        }
    }
})

$cmbHistory.Add_SelectionChanged({
    if ($cmbHistory.SelectedItem) {
        $txtPath.Text = $cmbHistory.SelectedItem
        Load-FolderTree $cmbHistory.SelectedItem
    }
})

$treeView.Add_SelectedItemChanged({
    if ($_.NewValue) {
        $selectedPath = $_.NewValue.Tag
        $script:currentPath = $selectedPath
        Load-FileList $selectedPath
        
        if ($_.NewValue.Items.Count -eq 1 -and $_.NewValue.Items[0].Header -eq "読み込み中...") {
            $_.NewValue.Items.Clear()
            Load-SubFolders -parentItem $_.NewValue -path $selectedPath -depth 0 -maxDepth 1
        }
    }
})

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

# 検索機能（再帰）
$txtSearch.Add_TextChanged({
    try {
        $searchText = $txtSearch.Text.Trim().ToLower()
        $dataGrid.Items.Clear()
        
        if (!$script:currentPath -or !(Test-Path $script:currentPath)) {
            return
        }
        
        if (![string]::IsNullOrWhiteSpace($searchText)) {
            $txtStatus.Text = "検索中..."
            
            $folders = Get-ChildItem -Path $script:currentPath -Directory -Recurse -ErrorAction SilentlyContinue |
                       Where-Object { $_.Name.ToLower().Contains($searchText) } |
                       Sort-Object FullName
            
            $files = Get-ChildItem -Path $script:currentPath -File -Recurse -ErrorAction SilentlyContinue |
                     Where-Object { $_.Name.ToLower().Contains($searchText) } |
                     Sort-Object FullName
            
            foreach ($folder in $folders) {
                $item = New-Object FileItem
                $item.Type = "フォルダ"
                $relativePath = $folder.FullName.Replace($script:currentPath, "").TrimStart('\')
                $item.Name = $relativePath
                $item.Size = "-"
                $item.Modified = $folder.LastWriteTime.ToString("yyyy/MM/dd HH:mm")
                $item.FullPath = $folder.FullName
                $dataGrid.Items.Add($item)
            }
            
            foreach ($file in $files) {
                $item = New-Object FileItem
                $item.Type = "ファイル"
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
            
            $totalCount = $folders.Count + $files.Count
            $txtStatus.Text = "検索結果（再帰的）: $totalCount 件（フォルダ: $($folders.Count) / ファイル: $($files.Count)）"
        }
        else {
            Load-FileList $script:currentPath
        }
    }
    catch {
        Write-Host "検索エラー: $_" -ForegroundColor Red
        $txtStatus.Text = "検索エラーが発生しました"
    }
})

$btnSearchClear.Add_Click({
    $txtSearch.Text = ""
    if ($script:currentPath) {
        Load-FileList $script:currentPath
    }
})

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

$btnNewFolder.Add_Click({
    Create-NewFolder
})

$btnTestFolder.Add_Click({
    Create-TestFolder
})

$btnCopy.Add_Click({
    Copy-SelectedItems $false
})

$btnCut.Add_Click({
    Copy-SelectedItems $true
})

$btnPaste.Add_Click({
    Paste-Items
})

$btnMoveToFolder.Add_Click({
    Write-Host "フォルダに移動機能を起動..." -ForegroundColor Yellow
    Move-ToSelectedFolder
})

$btnExportExcel.Add_Click({
    Write-Host "Excel出力機能を起動..." -ForegroundColor Yellow
    Export-FolderStructure $script:currentPath
})

$btnDelete.Add_Click({
    Write-Host "削除機能を起動..." -ForegroundColor Yellow
    Delete-SelectedItems
})

$btnBatchRename.Add_Click({
    Write-Host "一括リネーム機能を起動..." -ForegroundColor Yellow
    Show-BatchRenameDialog $script:currentPath
})

# 右クリックメニュー
$menuCopy.Add_Click({
    Copy-SelectedItems $false
})

$menuCut.Add_Click({
    Copy-SelectedItems $true
})

$menuPaste.Add_Click({
    Paste-Items
})

$menuDelete.Add_Click({
    Delete-SelectedItems
})

# キーボードショートカット
$dataGrid.Add_KeyDown({
    if ($_.Key -eq "Delete") {
        Delete-SelectedItems
    }
    elseif ($_.Key -eq "C" -and $_.KeyboardDevice.Modifiers -eq "Control") {
        Copy-SelectedItems $false
    }
    elseif ($_.Key -eq "X" -and $_.KeyboardDevice.Modifiers -eq "Control") {
        Copy-SelectedItems $true
    }
    elseif ($_.Key -eq "V" -and $_.KeyboardDevice.Modifiers -eq "Control") {
        Paste-Items
    }
})

# ===============================================
# 初期化
# ===============================================

# ブックマーク読み込み
Load-Bookmarks

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
Write-Host "  ✅ ブックマーク機能" -ForegroundColor White
Write-Host "  ✅ Excel/CSV出力機能" -ForegroundColor White
Write-Host "  ✅ ファイル削除機能" -ForegroundColor White
Write-Host "  ✅ ファイル移動/コピー機能" -ForegroundColor Green
Write-Host "  ✅ 新規フォルダ作成" -ForegroundColor White
Write-Host "  ✅ 試験項目フォルダ（連番）" -ForegroundColor White
Write-Host "  ✅ 一括リネーム機能（実装済み）" -ForegroundColor White
Write-Host "  ✅ ドラッグ&ドロップ（Excel/Text）" -ForegroundColor White
Write-Host "  ✅ 再帰検索機能（配下すべて）" -ForegroundColor White
Write-Host "  ✅ 履歴機能" -ForegroundColor White
Write-Host ""

# メインウィンドウ表示
$mainWindow.ShowDialog() | Out-Null

# 終了時にブックマーク保存
Save-Bookmarks

Write-Host "アプリケーションを終了しました" -ForegroundColor Yellow

