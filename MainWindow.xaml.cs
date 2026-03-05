using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Data;
using Microsoft.Win32;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace ExcelFinder;

public partial class MainWindow : Window
{
    private readonly ObservableCollection<JsonFileItem> _jsonFiles = [];
    private readonly ObservableCollection<UidMatchResult> _searchResults = [];
    private readonly AppSettingsStore _settingsStore;
    private readonly ICollectionView _jsonFilesView;
    private readonly ICollectionView _searchResultsView;
    private double? _savedJsonListHeight;
    private double? _savedWindowWidth;
    private double? _savedWindowHeight;

    public MainWindow()
    {
        InitializeComponent();

        _jsonFilesView = CollectionViewSource.GetDefaultView(_jsonFiles);
        _jsonFilesView.Filter = FilterJsonFile;
        JsonFilesListBox.ItemsSource = _jsonFilesView;
        _searchResultsView = CollectionViewSource.GetDefaultView(_searchResults);
        _searchResultsView.Filter = FilterSearchResult;
        ResultDataGrid.ItemsSource = _searchResultsView;

        _settingsStore = AppSettingsStore.CreateDefault();

        SetVersionText();
        LoadSavedFolders();
        RefreshJsonList();
        Loaded += MainWindow_Loaded;
        Closing += MainWindow_Closing;
    }

    private void SetVersionText()
    {
        Version? version = Assembly.GetExecutingAssembly().GetName().Version;
        int displayVersion = version?.Revision > 0 ? version.Revision : 1;
        string versionText = $"v{displayVersion}";
        VersionTextBlock.Text = versionText;
    }

    private void LoadSavedFolders()
    {
        var settings = _settingsStore.Load();

        ExcelFolderTextBox.Text = settings.ExcelFolder ?? string.Empty;
        JsonFolderTextBox.Text = settings.JsonFolder ?? string.Empty;
        _savedJsonListHeight = settings.JsonListHeight > 0 ? settings.JsonListHeight : null;
        _savedWindowWidth = settings.WindowWidth > 0 ? settings.WindowWidth : null;
        _savedWindowHeight = settings.WindowHeight > 0 ? settings.WindowHeight : null;

        ApplySavedWindowSize();
    }

    private void FolderPathTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        SaveFolders();

        if (ReferenceEquals(sender, JsonFolderTextBox))
        {
            RefreshJsonList();
        }
    }

    private void SaveFolders()
    {
        AppSettings settings = _settingsStore.Load();
        settings.ExcelFolder = ExcelFolderTextBox.Text.Trim();
        settings.JsonFolder = JsonFolderTextBox.Text.Trim();
        _settingsStore.Save(settings);
    }

    private void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        Dispatcher.BeginInvoke(ApplySavedJsonListHeight, DispatcherPriority.Loaded);
    }

    private void MainWindow_Closing(object? sender, CancelEventArgs e)
    {
        SaveJsonListHeightSetting();
        SaveWindowSizeSetting();
    }

    private void BrowseExcelFolderButton_Click(object sender, RoutedEventArgs e)
    {
        string? selectedPath = SelectFolder("Excel 폴더 선택", ExcelFolderTextBox.Text);
        if (selectedPath is null)
        {
            return;
        }

        ExcelFolderTextBox.Text = selectedPath;
    }

    private void BrowseJsonFolderButton_Click(object sender, RoutedEventArgs e)
    {
        string? selectedPath = SelectFolder("JSON 폴더 선택", JsonFolderTextBox.Text);
        if (selectedPath is null)
        {
            return;
        }

        JsonFolderTextBox.Text = selectedPath;
        RefreshJsonList();
    }

    private string? SelectFolder(string title, string currentPath)
    {
        AppSettings settings = _settingsStore.Load();
        string lastBrowseFolder = settings.LastBrowseFolder ?? string.Empty;
        string fallbackDirectory = Directory.Exists(lastBrowseFolder)
            ? lastBrowseFolder
            : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        var dialog = new OpenFolderDialog
        {
            Title = title,
            InitialDirectory = Directory.Exists(currentPath) ? currentPath : fallbackDirectory
        };

        bool? result = dialog.ShowDialog(this);
        if (result != true)
        {
            return null;
        }

        SaveLastBrowseFolder(dialog.FolderName);
        return dialog.FolderName;
    }

    private void SaveLastBrowseFolder(string folderPath)
    {
        if (string.IsNullOrWhiteSpace(folderPath))
        {
            return;
        }

        AppSettings settings = _settingsStore.Load();
        settings.LastBrowseFolder = folderPath.Trim();
        _settingsStore.Save(settings);
    }

    private async void RefreshJsonListButton_Click(object sender, RoutedEventArgs e)
    {
        RefreshJsonList();
        await RunSearchAsync(runWhenUidEmpty: false);
    }

    private void RefreshJsonList()
    {
        _jsonFiles.Clear();

        string jsonFolder = JsonFolderTextBox.Text.Trim();
        if (!Directory.Exists(jsonFolder))
        {
            StatusTextBlock.Text = "JSON 폴더 경로를 확인해 주세요.";
            return;
        }

        foreach (string filePath in Directory.EnumerateFiles(jsonFolder, "*.json", SearchOption.TopDirectoryOnly)
                     .OrderBy(Path.GetFileName))
        {
            _jsonFiles.Add(new JsonFileItem(filePath));
        }

        _jsonFilesView.Refresh();

        if (_jsonFiles.Count > 0)
        {
            if (_jsonFilesView.Cast<JsonFileItem>().Any())
            {
                JsonFilesListBox.SelectedIndex = 0;
            }
            StatusTextBlock.Text = $"JSON 파일 {_jsonFiles.Count}개를 불러왔습니다.";
        }
        else
        {
            StatusTextBlock.Text = "JSON 폴더에 파일이 없습니다.";
        }
    }

    private void JsonFilterTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        _jsonFilesView.Refresh();
        if (_jsonFilesView.Cast<JsonFileItem>().Any() && JsonFilesListBox.SelectedItem is null)
        {
            JsonFilesListBox.SelectedIndex = 0;
        }
    }

    private bool FilterJsonFile(object item)
    {
        if (item is not JsonFileItem file)
        {
            return false;
        }

        string keyword = JsonFilterTextBox?.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(keyword))
        {
            return true;
        }

        return file.DisplayName.Contains(keyword, StringComparison.OrdinalIgnoreCase);
    }

    private void ResultFilterTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        _searchResultsView.Refresh();
        if (_searchResultsView.Cast<UidMatchResult>().Any() && ResultDataGrid.SelectedItem is null)
        {
            ResultDataGrid.SelectedIndex = 0;
        }
    }

    private bool FilterSearchResult(object item)
    {
        if (item is not UidMatchResult result)
        {
            return false;
        }

        string keyword = ResultFilterTextBox?.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(keyword))
        {
            return true;
        }

        return result.FileName.Contains(keyword, StringComparison.OrdinalIgnoreCase)
               || result.SheetName.Contains(keyword, StringComparison.OrdinalIgnoreCase)
               || result.RowContent.Contains(keyword, StringComparison.OrdinalIgnoreCase)
               || result.RowNumber.ToString().Contains(keyword, StringComparison.OrdinalIgnoreCase);
    }

    private void JsonFilesListBox_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (JsonFilesListBox.SelectedItem is not JsonFileItem selectedJson)
        {
            return;
        }

        if (!File.Exists(selectedJson.FullPath))
        {
            StatusTextBlock.Text = $"파일을 찾을 수 없습니다: {selectedJson.FullPath}";
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = selectedJson.FullPath,
                UseShellExecute = true
            });
            StatusTextBlock.Text = $"열기: {selectedJson.DisplayName}";
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"파일 열기 실패: {ex.Message}";
        }
    }

    private async void SearchButton_Click(object sender, RoutedEventArgs e)
    {
        await RunSearchAsync(runWhenUidEmpty: false);
    }

    private async Task RunSearchAsync(bool runWhenUidEmpty)
    {
        _searchResults.Clear();
        _searchResultsView.Refresh();

        string uid = UidTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(uid) && !runWhenUidEmpty)
        {
            StatusTextBlock.Text = "UID를 입력해 주세요.";
            return;
        }

        if (JsonFilesListBox.SelectedItem is not JsonFileItem selectedJson)
        {
            StatusTextBlock.Text = "JSON 파일을 선택해 주세요.";
            return;
        }

        string excelFolder = ExcelFolderTextBox.Text.Trim();
        if (!Directory.Exists(excelFolder))
        {
            StatusTextBlock.Text = "Excel 폴더 경로를 확인해 주세요.";
            return;
        }

        SaveFolders();

        SearchButton.IsEnabled = false;
        StatusTextBlock.Text = $"검색 중... (기준 JSON: {selectedJson.DisplayName})";

        try
        {
            bool includeAllContent = IncludeAllContentCheckBox.IsChecked == true;
            List<UidMatchResult> matches = await Task.Run(() => FindExcelFilesContainingUid(excelFolder, uid, includeAllContent));

            foreach (UidMatchResult match in matches)
            {
                _searchResults.Add(match);
            }
            _searchResultsView.Refresh();

            int fileCount = matches.Select(x => x.FilePath).Distinct(StringComparer.OrdinalIgnoreCase).Count();
            StatusTextBlock.Text = matches.Count > 0
                ? $"완료: 파일 {fileCount}개 / 행 {matches.Count}건"
                : includeAllContent
                    ? $"완료: UID '{uid}'를 포함한 Excel 파일이 없습니다."
                    : $"완료: UID '{uid}'와 정확히 일치하는 값이 없습니다.";
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"오류: {ex.Message}";
        }
        finally
        {
            SearchButton.IsEnabled = true;
        }
    }

    private void ResultDataGrid_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (ResultDataGrid.SelectedItem is not UidMatchResult selectedItem)
        {
            return;
        }

        if (!File.Exists(selectedItem.FilePath))
        {
            StatusTextBlock.Text = $"파일을 찾을 수 없습니다: {selectedItem.FilePath}";
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = selectedItem.FilePath,
                UseShellExecute = true
            });
            StatusTextBlock.Text = $"열기: {selectedItem.FileName}";
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"파일 열기 실패: {ex.Message}";
        }
    }

    private void ResultDataGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        DependencyObject? source = e.OriginalSource as DependencyObject;
        if (source is null)
        {
            return;
        }

        DataGridRow? row = FindVisualParent<DataGridRow>(source);
        if (row?.Item is UidMatchResult)
        {
            row.IsSelected = true;
            ResultDataGrid.SelectedItem = row.Item;
        }
    }

    private void ResultContextMenu_Opened(object sender, RoutedEventArgs e)
    {
        if (ResultDataGrid.SelectedItem is UidMatchResult selectedItem && selectedItem.IsCheckedOut)
        {
            WhoMenuItem.Visibility = Visibility.Visible;
            return;
        }

        WhoMenuItem.Visibility = Visibility.Collapsed;
    }

    private void OpenContainingFolderMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (ResultDataGrid.SelectedItem is not UidMatchResult selectedItem)
        {
            StatusTextBlock.Text = "선택된 검색 결과가 없습니다.";
            return;
        }

        if (!File.Exists(selectedItem.FilePath))
        {
            StatusTextBlock.Text = $"파일을 찾을 수 없습니다: {selectedItem.FilePath}";
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"/select,\"{selectedItem.FilePath}\"",
                UseShellExecute = true
            });
            StatusTextBlock.Text = $"탐색기 열기: {selectedItem.FileName}";
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"탐색기 열기 실패: {ex.Message}";
        }
    }

    private async void PerforceCheckoutMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (ResultDataGrid.SelectedItem is not UidMatchResult selectedItem)
        {
            StatusTextBlock.Text = "선택된 검색 결과가 없습니다.";
            return;
        }

        if (!File.Exists(selectedItem.FilePath))
        {
            StatusTextBlock.Text = $"파일을 찾을 수 없습니다: {selectedItem.FilePath}";
            return;
        }

        StatusTextBlock.Text = $"Perforce 체크아웃 중: {selectedItem.FileName}";

        try
        {
            (bool success, string message) = await Task.Run(() => PerforceHelper.Checkout(selectedItem.FilePath));
            if (success)
            {
                RefreshCheckoutStateInSearchResults(selectedItem.FilePath, true);
                StatusTextBlock.Text = $"Perforce 체크아웃 완료: {selectedItem.FileName}";
                return;
            }

            StatusTextBlock.Text = $"Perforce 체크아웃 실패: {message}";
            OpenPerforceConfigWindow();
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"Perforce 실행 오류: {ex.Message}";
        }
    }

    private async void WhoMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (ResultDataGrid.SelectedItem is not UidMatchResult selectedItem)
        {
            StatusTextBlock.Text = "선택된 검색 결과가 없습니다.";
            return;
        }

        if (!selectedItem.IsCheckedOut)
        {
            MessageBox.Show(this, "체크아웃 상태가 아닙니다.", "Who?", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        try
        {
            (bool success, string message, List<string> clients) = await Task.Run(() => PerforceHelper.GetOpenedClientWorkspaces(selectedItem.FilePath));
            if (!success)
            {
                MessageBox.Show(this, $"조회 실패: {message}", "Who?", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (clients.Count == 0)
            {
                MessageBox.Show(this, "체크아웃된 Client Workspace 정보를 찾지 못했습니다.", "Who?", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string text = string.Join(Environment.NewLine, clients.Select((x, i) => $"{i + 1}. {x}"));
            MessageBox.Show(this, $"Client Workspace\n{text}", "Who?", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"조회 중 오류: {ex.Message}", "Who?", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void RefreshCheckoutStateInSearchResults(string filePath, bool isCheckedOut)
    {
        bool changed = false;
        foreach (UidMatchResult item in _searchResults)
        {
            if (!string.Equals(item.FilePath, filePath, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            item.IsCheckedOut = isCheckedOut;
            changed = true;
        }

        if (changed)
        {
            _searchResultsView.Refresh();
            ResultDataGrid.Items.Refresh();
        }
    }

    private void OpenPerforceConfigMenuItem_Click(object sender, RoutedEventArgs e)
    {
        OpenPerforceConfigWindow();
    }

    private void OpenMergeWindowMenuItem_Click(object sender, RoutedEventArgs e)
    {
        var mergeWindow = new MergeWindow
        {
            Owner = this
        };
        mergeWindow.ShowDialog();
    }

    private void OpenJsonExtractWindowMenuItem_Click(object sender, RoutedEventArgs e)
    {
        string excelFolder = ExcelFolderTextBox.Text.Trim();
        if (!Directory.Exists(excelFolder))
        {
            StatusTextBlock.Text = "Excel 폴더 경로를 확인해 주세요.";
            return;
        }

        string jsonFolder = JsonFolderTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(jsonFolder))
        {
            StatusTextBlock.Text = "JSON 폴더 경로를 확인해 주세요.";
            return;
        }

        var extractWindow = new JsonExtractWindow(excelFolder, jsonFolder)
        {
            Owner = this
        };

        bool? result = extractWindow.ShowDialog();
        if (result == true)
        {
            RefreshJsonList();
            StatusTextBlock.Text = "JsonExporter 작업이 완료되었습니다.";
        }
    }

    private void OpenJsonDiffWindowMenuItem_Click(object sender, RoutedEventArgs e)
    {
        var diffWindow = new JsonDiffWindow
        {
            Owner = this
        };
        diffWindow.ShowDialog();
    }

    private void OpenPerforceConfigWindow()
    {
        PerforceClientSettings? initialSettings = null;
        (bool success, string _, PerforceClientSettings? settings) = PerforceHelper.GetClientSettings();
        if (success)
        {
            initialSettings = settings;
        }

        var configWindow = new PerforceConfigWindow(initialSettings)
        {
            Owner = this
        };

        bool? result = configWindow.ShowDialog();
        if (result == true)
        {
            StatusTextBlock.Text = "Perforce 정보가 적용되었습니다.";
        }
    }

    private void OpenEditorMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (ResultDataGrid.SelectedItem is not UidMatchResult selectedItem)
        {
            StatusTextBlock.Text = "선택된 검색 결과가 없습니다.";
            return;
        }

        if (!File.Exists(selectedItem.FilePath))
        {
            StatusTextBlock.Text = $"파일을 찾을 수 없습니다: {selectedItem.FilePath}";
            return;
        }

        var editor = new ExcelEditorWindow(selectedItem.FilePath, selectedItem.SheetName, selectedItem.RowNumber)
        {
            Owner = this
        };
        editor.Show();
    }

    private void OpenHistoryMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (ResultDataGrid.SelectedItem is not UidMatchResult selectedItem)
        {
            StatusTextBlock.Text = "선택된 검색 결과가 없습니다.";
            return;
        }

        if (!File.Exists(selectedItem.FilePath))
        {
            StatusTextBlock.Text = $"파일을 찾을 수 없습니다: {selectedItem.FilePath}";
            return;
        }

        var historyWindow = new HistoryWindow(selectedItem.FilePath)
        {
            Owner = this
        };
        historyWindow.ShowDialog();
    }

    private void JsonResultGridSplitter_DragCompleted(object sender, DragCompletedEventArgs e)
    {
        SaveJsonListHeightSetting();
    }

    private void ApplySavedJsonListHeight()
    {
        if (!_savedJsonListHeight.HasValue)
        {
            return;
        }

        double height = _savedJsonListHeight.Value;
        if (double.IsNaN(height) || height < 60)
        {
            return;
        }

        JsonListRowDefinition.Height = new GridLength(height, GridUnitType.Pixel);
        ResultListRowDefinition.Height = new GridLength(1, GridUnitType.Star);
    }

    private void SaveJsonListHeightSetting()
    {
        double currentHeight = JsonListRowDefinition.ActualHeight;
        if (currentHeight < 60 || double.IsNaN(currentHeight))
        {
            return;
        }

        AppSettings settings = _settingsStore.Load();
        settings.JsonListHeight = currentHeight;
        _settingsStore.Save(settings);
    }

    private void ApplySavedWindowSize()
    {
        if (_savedWindowWidth.HasValue && _savedWindowWidth.Value >= 700)
        {
            Width = _savedWindowWidth.Value;
        }

        if (_savedWindowHeight.HasValue && _savedWindowHeight.Value >= 500)
        {
            Height = _savedWindowHeight.Value;
        }
    }

    private void SaveWindowSizeSetting()
    {
        Rect bounds = RestoreBounds;
        if (bounds.Width < 700 || bounds.Height < 500)
        {
            return;
        }

        AppSettings settings = _settingsStore.Load();
        settings.WindowWidth = bounds.Width;
        settings.WindowHeight = bounds.Height;
        _settingsStore.Save(settings);
    }

    private static T? FindVisualParent<T>(DependencyObject? child) where T : DependencyObject
    {
        while (child is not null)
        {
            if (child is T typed)
            {
                return typed;
            }

            child = VisualTreeHelper.GetParent(child);
        }

        return null;
    }

    private static List<UidMatchResult> FindExcelFilesContainingUid(string excelFolder, string uid, bool includeAllContent)
    {
        var matches = new List<UidMatchResult>();
        var checkoutStatusCache = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

        IEnumerable<string> excelFiles = Directory.EnumerateFiles(excelFolder, "*.*", SearchOption.AllDirectories)
            .Where(file =>
            {
                string ext = Path.GetExtension(file);
                return ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
                       || ext.Equals(".xlsm", StringComparison.OrdinalIgnoreCase);
            });

        foreach (string file in excelFiles)
        {
            List<UidMatchResult> fileMatches = FindUidRowsInExcel(file, uid, includeAllContent).ToList();
            if (fileMatches.Count == 0)
            {
                continue;
            }

            bool isCheckedOut = GetCheckedOutStatusCached(file, checkoutStatusCache);
            foreach (UidMatchResult match in fileMatches)
            {
                match.IsCheckedOut = isCheckedOut;
                matches.Add(match);
            }
        }

        return matches;
    }

    private static IEnumerable<UidMatchResult> FindUidRowsInExcel(string excelFilePath, string uid, bool includeAllContent)
    {
        var matches = new List<UidMatchResult>();
        string uidTrimmed = uid.Trim();
        if (string.IsNullOrEmpty(uidTrimmed))
        {
            return matches;
        }

        try
        {
            using var stream = new FileStream(
                excelFilePath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            List<string> sharedStrings = LoadSharedStrings(archive);
            Dictionary<string, string> worksheetNameMap = LoadWorksheetNameMap(archive);

            foreach (ZipArchiveEntry worksheet in archive.Entries.Where(e => e.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)
                                                                             && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
            {
                using Stream worksheetStream = worksheet.Open();
                XDocument doc = XDocument.Load(worksheetStream);
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                int? uidColumnIndex = null;

                foreach (XElement row in doc.Descendants(ns + "row"))
                {
                    var rowTexts = new List<string>();
                    var rowCellValuesByColumn = new Dictionary<int, string>();
                    bool containsUid = false;

                    foreach (XElement cell in row.Elements(ns + "c"))
                    {
                        string cellText = ResolveCellText(cell, sharedStrings, ns);
                        if (string.IsNullOrWhiteSpace(cellText))
                        {
                            continue;
                        }

                        rowTexts.Add(cellText.Trim());
                        int? columnIndex = GetColumnIndexFromCell(cell);
                        if (columnIndex.HasValue)
                        {
                            rowCellValuesByColumn[columnIndex.Value] = cellText.Trim();
                        }

                        if (!containsUid && includeAllContent && IsCellMatch(cellText, uidTrimmed, includeAllContent))
                        {
                            containsUid = true;
                        }
                    }

                    // "모든 내용 포함 검색"이 꺼진 경우에는 헤더에서 UID 컬럼을 찾고 해당 컬럼만 비교한다.
                    if (!includeAllContent)
                    {
                        if (!uidColumnIndex.HasValue)
                        {
                            uidColumnIndex = FindUidColumnIndex(rowCellValuesByColumn);
                            continue;
                        }

                        if (rowCellValuesByColumn.TryGetValue(uidColumnIndex.Value, out string? uidCellText)
                            && IsCellMatch(uidCellText, uidTrimmed, includeAllContent))
                        {
                            containsUid = true;
                        }
                    }

                    if (!containsUid)
                    {
                        continue;
                    }

                    int rowNumber = ParseRowNumber(row);
                    string rowContent = string.Join(" | ", rowTexts);
                    if (string.IsNullOrWhiteSpace(rowContent))
                    {
                        rowContent = "(빈 행)";
                    }

                    string sheetName = worksheetNameMap.TryGetValue(worksheet.FullName, out string? name)
                        ? name
                        : Path.GetFileNameWithoutExtension(worksheet.Name);

                    matches.Add(new UidMatchResult
                    {
                        FilePath = excelFilePath,
                        SheetName = sheetName,
                        RowNumber = rowNumber,
                        RowContent = rowContent
                    });
                }
            }
        }
        catch
        {
            // 단일 파일 오류는 전체 검색을 중단하지 않기 위해 무시
        }

        return matches;
    }

    private static bool GetCheckedOutStatusCached(string filePath, Dictionary<string, bool> cache)
    {
        if (cache.TryGetValue(filePath, out bool cached))
        {
            return cached;
        }

        try
        {
            (bool success, bool isCheckedOut, _) = PerforceHelper.GetCheckoutStatus(filePath);
            bool result = success && isCheckedOut;
            cache[filePath] = result;
            return result;
        }
        catch
        {
            cache[filePath] = false;
            return false;
        }
    }

    private static int? FindUidColumnIndex(Dictionary<int, string> rowCellValuesByColumn)
    {
        foreach ((int columnIndex, string value) in rowCellValuesByColumn)
        {
            if (string.Equals(value.Trim(), "UID", StringComparison.OrdinalIgnoreCase))
            {
                return columnIndex;
            }
        }

        return null;
    }

    private static bool IsCellMatch(string cellText, string uid, bool includeAllContent)
    {
        if (includeAllContent)
        {
            return cellText.Contains(uid, StringComparison.OrdinalIgnoreCase);
        }

        return string.Equals(cellText.Trim(), uid, StringComparison.OrdinalIgnoreCase);
    }

    private static List<string> LoadSharedStrings(ZipArchive archive)
    {
        var result = new List<string>();
        ZipArchiveEntry? entry = archive.GetEntry("xl/sharedStrings.xml");
        if (entry is null)
        {
            return result;
        }

        using Stream stream = entry.Open();
        XDocument doc = XDocument.Load(stream);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        foreach (XElement item in doc.Descendants(ns + "si"))
        {
            IEnumerable<string> textParts = item.Descendants(ns + "t").Select(t => t.Value);
            result.Add(string.Concat(textParts));
        }

        return result;
    }

    private static Dictionary<string, string> LoadWorksheetNameMap(ZipArchive archive)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        ZipArchiveEntry? workbookEntry = archive.GetEntry("xl/workbook.xml");
        ZipArchiveEntry? relsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
        if (workbookEntry is null || relsEntry is null)
        {
            return result;
        }

        using Stream relsStream = relsEntry.Open();
        XDocument relsDoc = XDocument.Load(relsStream);
        XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";

        var relMap = relsDoc.Descendants(relNs + "Relationship")
            .Select(x => new
            {
                Id = (string?)x.Attribute("Id"),
                Target = (string?)x.Attribute("Target")
            })
            .Where(x => !string.IsNullOrWhiteSpace(x.Id) && !string.IsNullOrWhiteSpace(x.Target))
            .ToDictionary(x => x.Id!, x => NormalizeWorksheetPath(x.Target!), StringComparer.OrdinalIgnoreCase);

        using Stream workbookStream = workbookEntry.Open();
        XDocument workbookDoc = XDocument.Load(workbookStream);
        XNamespace wbNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace relAttrNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        foreach (XElement sheet in workbookDoc.Descendants(wbNs + "sheet"))
        {
            string? relId = (string?)sheet.Attribute(relAttrNs + "id");
            string? sheetName = (string?)sheet.Attribute("name");
            if (string.IsNullOrWhiteSpace(relId) || string.IsNullOrWhiteSpace(sheetName))
            {
                continue;
            }

            if (relMap.TryGetValue(relId, out string? worksheetPath))
            {
                result[worksheetPath] = sheetName;
            }
        }

        return result;
    }

    private static string NormalizeWorksheetPath(string target)
    {
        string normalized = target.Replace('\\', '/').TrimStart('/');
        if (!normalized.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
        {
            normalized = $"xl/{normalized}";
        }

        return normalized;
    }

    private static string ResolveCellText(XElement cell, List<string> sharedStrings, XNamespace ns)
    {
        string cellType = (string?)cell.Attribute("t") ?? string.Empty;

        if (string.Equals(cellType, "s", StringComparison.OrdinalIgnoreCase))
        {
            string indexText = cell.Element(ns + "v")?.Value ?? string.Empty;
            if (int.TryParse(indexText, out int idx) && idx >= 0 && idx < sharedStrings.Count)
            {
                return sharedStrings[idx];
            }
        }

        if (string.Equals(cellType, "inlineStr", StringComparison.OrdinalIgnoreCase))
        {
            IEnumerable<string> inlineParts = cell.Descendants(ns + "t").Select(t => t.Value);
            return string.Concat(inlineParts);
        }

        return cell.Element(ns + "v")?.Value ?? string.Empty;
    }

    private static int ParseRowNumber(XElement row)
    {
        string rowValue = (string?)row.Attribute("r") ?? string.Empty;
        return int.TryParse(rowValue, out int rowNumber) ? rowNumber : -1;
    }

    private static int? GetColumnIndexFromCell(XElement cell)
    {
        string reference = (string?)cell.Attribute("r") ?? string.Empty;
        if (string.IsNullOrWhiteSpace(reference))
        {
            return null;
        }

        Match m = Regex.Match(reference, "^[A-Za-z]+");
        if (!m.Success)
        {
            return null;
        }

        string letters = m.Value.ToUpperInvariant();
        int columnIndex = 0;
        foreach (char c in letters)
        {
            columnIndex = (columnIndex * 26) + (c - 'A' + 1);
        }

        return columnIndex;
    }
}

public sealed class JsonFileItem
{
    public JsonFileItem(string fullPath)
    {
        FullPath = fullPath;
        DisplayName = Path.GetFileName(fullPath);
    }

    public string DisplayName { get; }
    public string FullPath { get; }
}

public sealed class AppSettings
{
    public string ExcelFolder { get; set; } = string.Empty;
    public string JsonFolder { get; set; } = string.Empty;
    public string JsonExporterExcelFolder { get; set; } = string.Empty;
    public string JsonExporterJsonFolder { get; set; } = string.Empty;
    public List<string> JsonExporterCheckedNames { get; set; } = [];
    public string JsonDiffSourceFilePath { get; set; } = string.Empty;
    public string JsonDiffTargetFilePath { get; set; } = string.Empty;
    public string JsonDiffSourceDirPath { get; set; } = string.Empty;
    public string JsonDiffTargetDirPath { get; set; } = string.Empty;
    public string CheckinPrefix { get; set; } = string.Empty;
    public double JsonListHeight { get; set; }
    public string LastBrowseFolder { get; set; } = string.Empty;
    public double WindowWidth { get; set; }
    public double WindowHeight { get; set; }
}

public sealed class UidMatchResult
{
    public string FilePath { get; set; } = string.Empty;
    public string FileName => Path.GetFileName(FilePath);
    public string SheetName { get; set; } = string.Empty;
    public int RowNumber { get; set; }
    public string RowContent { get; set; } = string.Empty;
    public bool IsCheckedOut { get; set; }
}

public sealed class AppSettingsStore
{
    private readonly string _settingsFilePath;

    private AppSettingsStore(string settingsFilePath)
    {
        _settingsFilePath = settingsFilePath;
    }

    public static AppSettingsStore CreateDefault()
    {
        string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        string appDir = Path.Combine(appData, "ExcelFinder");
        string filePath = Path.Combine(appDir, "settings.json");
        return new AppSettingsStore(filePath);
    }

    public AppSettings Load()
    {
        try
        {
            if (!File.Exists(_settingsFilePath))
            {
                return new AppSettings();
            }

            string json = File.ReadAllText(_settingsFilePath);
            return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
        }
        catch
        {
            return new AppSettings();
        }
    }

    public void Save(AppSettings settings)
    {
        try
        {
            string? directory = Path.GetDirectoryName(_settingsFilePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            string json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_settingsFilePath, json);
        }
        catch
        {
            // 설정 저장 실패는 앱 동작에 치명적이지 않으므로 무시
        }
    }
}
