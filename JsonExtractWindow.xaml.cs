using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace ExcelFinder;

public partial class JsonExtractWindow : Window
{
    private readonly AppSettingsStore _settingsStore = AppSettingsStore.CreateDefault();
    private readonly HashSet<string> _restoredCheckedNames = new(StringComparer.OrdinalIgnoreCase);
    private string _excelFolder;
    private string _jsonFolder;
    private readonly ObservableCollection<JsonExtractNameItem> _items = [];
    private bool _nextToggleState = true;
    private bool _hasExtracted;
    private bool _suppressPathSave;
    private bool _suppressCheckedSave;
    private string _contextCellValue = string.Empty;

    public JsonExtractWindow(string excelFolder, string jsonFolder)
    {
        InitializeComponent();

        AppSettings settings = _settingsStore.Load();
        _excelFolder = !string.IsNullOrWhiteSpace(settings.JsonExporterExcelFolder)
            ? settings.JsonExporterExcelFolder
            : excelFolder;
        _jsonFolder = !string.IsNullOrWhiteSpace(settings.JsonExporterJsonFolder)
            ? settings.JsonExporterJsonFolder
            : jsonFolder;
        foreach (string name in settings.JsonExporterCheckedNames ?? [])
        {
            if (!string.IsNullOrWhiteSpace(name))
            {
                _restoredCheckedNames.Add(name.Trim());
            }
        }

        _suppressPathSave = true;
        _suppressCheckedSave = true;
        ItemsDataGrid.ItemsSource = _items;
        ExcelPathTextBox.Text = _excelFolder;
        JsonPathTextBox.Text = _jsonFolder;
        _suppressPathSave = false;
        _suppressCheckedSave = false;

        SaveExporterPaths();
        LoadCandidates();
    }

    private void LoadCandidates()
    {
        _items.Clear();

        if (!Directory.Exists(_excelFolder))
        {
            StatusTextBlock.Text = "Excel 폴더 경로를 확인해 주세요.";
            return;
        }

        var allExcelFiles = Directory.EnumerateFiles(_excelFolder, "*.*", SearchOption.AllDirectories)
            .Where(IsExcelFile)
            .ToList();

        var candidates = new List<JsonExtractCandidate>();
        foreach (string path in allExcelFiles)
        {
            string stem = Path.GetFileNameWithoutExtension(path);
            if (!stem.StartsWith("Data_", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            string name = ExtractNameToken(stem);
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            string groupKey = ExtractGroupKey(stem);
            candidates.Add(new JsonExtractCandidate(path, stem, name, groupKey));
        }

        foreach (IGrouping<string, JsonExtractCandidate> byName in candidates
                     .GroupBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
                     .OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase))
        {
            List<JsonExtractCandidate> byNameItems = byName
                .OrderBy(x => x.GroupKey, StringComparer.OrdinalIgnoreCase)
                .ThenBy(x => x.FileStem, StringComparer.OrdinalIgnoreCase)
                .ToList();
            string groupSummary = string.Join(", ", byNameItems.Select(x => x.GroupKey)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase));

            var item = new JsonExtractNameItem
            {
                IsChecked = _restoredCheckedNames.Contains(byName.Key),
                Name = byName.Key,
                GroupSummary = groupSummary,
                Candidates = byNameItems
            };
            item.PropertyChanged += JsonExtractNameItem_PropertyChanged;
            _items.Add(item);
        }

        if (_items.Count == 0)
        {
            StatusTextBlock.Text = "JsonExporter 대상(Data_*.xlsx/.xlsm) 파일이 없습니다.";
            return;
        }

        int fileCount = _items.Sum(x => x.FileCount);
        StatusTextBlock.Text = $"후보 이름 {_items.Count}개 / 대상 Excel {fileCount}개";
    }

    private static bool IsExcelFile(string path)
    {
        string ext = Path.GetExtension(path);
        return ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
               || ext.Equals(".xlsm", StringComparison.OrdinalIgnoreCase);
    }

    private static string ExtractNameToken(string stem)
    {
        string rest = stem.StartsWith("Data_", StringComparison.OrdinalIgnoreCase)
            ? stem.Substring("Data_".Length)
            : stem;

        if (string.IsNullOrWhiteSpace(rest))
        {
            return string.Empty;
        }

        int idx = rest.IndexOf('_');
        return idx < 0 ? rest.Trim() : rest.Substring(0, idx).Trim();
    }

    private static string ExtractGroupKey(string stem)
    {
        string[] parts = stem.Split('_', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length >= 4 && string.Equals(parts[0], "Data", StringComparison.OrdinalIgnoreCase))
        {
            return string.Join("_", parts.Take(parts.Length - 1));
        }

        return stem;
    }

    private void ToggleAllButton_Click(object sender, RoutedEventArgs e)
    {
        bool setTo = _nextToggleState;
        foreach (JsonExtractNameItem item in _items)
        {
            item.IsChecked = setTo;
        }

        _nextToggleState = !_nextToggleState;
        SaveCheckedNames();
        StatusTextBlock.Text = setTo ? "전체 체크" : "전체 체크 해제";
        ItemsDataGrid.Items.Refresh();
    }

    private void ItemsDataGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        _contextCellValue = string.Empty;

        DependencyObject? source = e.OriginalSource as DependencyObject;
        if (source is null)
        {
            return;
        }

        DataGridCell? cell = FindVisualParent<DataGridCell>(source);
        DataGridRow? row = FindVisualParent<DataGridRow>(source);
        if (row?.Item is not JsonExtractNameItem item)
        {
            return;
        }

        row.IsSelected = true;
        ItemsDataGrid.SelectedItem = row.Item;
        int displayIndex = cell?.Column?.DisplayIndex ?? -1;
        _contextCellValue = GetItemCellValueByDisplayIndex(item, displayIndex);
    }

    private void CopyItemCellValueMenuItem_Click(object sender, RoutedEventArgs e)
    {
        string value = _contextCellValue;
        if (string.IsNullOrEmpty(value) && ItemsDataGrid.SelectedItem is JsonExtractNameItem selected)
        {
            value = selected.Name;
        }

        try
        {
            Clipboard.SetText(value ?? string.Empty);
            StatusTextBlock.Text = "값을 클립보드에 복사했습니다.";
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"복사 실패: {ex.Message}";
        }
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = _hasExtracted;
        Close();
    }

    private void ExcelPathTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        _excelFolder = ExcelPathTextBox.Text.Trim();
        SaveExporterPaths();
        LoadCandidates();
    }

    private void JsonPathTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        _jsonFolder = JsonPathTextBox.Text.Trim();
        SaveExporterPaths();
    }

    private void BrowseExcelPathButton_Click(object sender, RoutedEventArgs e)
    {
        string initialDir = Directory.Exists(ExcelPathTextBox.Text.Trim())
            ? ExcelPathTextBox.Text.Trim()
            : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var dialog = new OpenFolderDialog
        {
            Title = "Excel 폴더 선택",
            InitialDirectory = initialDir
        };

        bool? result = dialog.ShowDialog(this);
        if (result == true)
        {
            ExcelPathTextBox.Text = dialog.FolderName;
        }
    }

    private void BrowseJsonPathButton_Click(object sender, RoutedEventArgs e)
    {
        string initialDir = Directory.Exists(JsonPathTextBox.Text.Trim())
            ? JsonPathTextBox.Text.Trim()
            : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var dialog = new OpenFolderDialog
        {
            Title = "JSON 폴더 선택",
            InitialDirectory = initialDir
        };

        bool? result = dialog.ShowDialog(this);
        if (result == true)
        {
            JsonPathTextBox.Text = dialog.FolderName;
        }
    }

    private async void ExtractButton_Click(object sender, RoutedEventArgs e)
    {
        // 체크박스를 클릭한 직후 추출 버튼을 눌러도 최신 체크 상태를 반영한다.
        ItemsDataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
        ItemsDataGrid.CommitEdit(DataGridEditingUnit.Row, true);

        _jsonFolder = JsonPathTextBox.Text.Trim();
        SaveExporterPaths();
        if (string.IsNullOrWhiteSpace(_jsonFolder))
        {
            StatusTextBlock.Text = "JSON 폴더 경로를 확인해 주세요.";
            return;
        }

        List<JsonExtractNameItem> selectedNames = _items.Where(x => x.IsChecked).ToList();
        if (selectedNames.Count == 0)
        {
            StatusTextBlock.Text = "JsonExporter 할 항목을 체크해 주세요.";
            return;
        }

        Directory.CreateDirectory(_jsonFolder);

        var grouped = selectedNames
            .SelectMany(x => x.Candidates)
            .GroupBy(x => x.GroupKey, StringComparer.OrdinalIgnoreCase)
            .OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase)
            .ToList();

        SetExtractionUiBusy(true, grouped.Count);

        int created = 0;
        var createdFileNames = new List<string>();
        try
        {
            for (int i = 0; i < grouped.Count; i++)
            {
                IGrouping<string, JsonExtractCandidate> group = grouped[i];
                List<Dictionary<string, string>> rows = await Task.Run(() => BuildGroupRows(group.ToList()));
                string outputFileName = $"{group.Key}.json";
                string outputPath = Path.Combine(_jsonFolder, outputFileName);

                string json = JsonSerializer.Serialize(rows, new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                });

                await File.WriteAllTextAsync(outputPath, json);
                created++;
                createdFileNames.Add(outputFileName);
                UpdateExtractionProgress(i + 1, grouped.Count, outputFileName);
            }

            _hasExtracted = true;
            SaveCheckedNames();
            StatusTextBlock.Text = $"추출 완료: JSON {created}개";
            ShowExportResultWindow(createdFileNames);
        }
        finally
        {
            SetExtractionUiBusy(false, 0);
        }
    }

    private void SetExtractionUiBusy(bool isBusy, int total)
    {
        ToggleAllButton.IsEnabled = !isBusy;
        ExtractButton.IsEnabled = !isBusy;
        ItemsDataGrid.IsEnabled = !isBusy;
        ExcelPathTextBox.IsEnabled = !isBusy;
        JsonPathTextBox.IsEnabled = !isBusy;
        ProgressPanel.Visibility = isBusy ? Visibility.Visible : Visibility.Collapsed;

        if (isBusy)
        {
            ExtractProgressBar.Minimum = 0;
            ExtractProgressBar.Maximum = Math.Max(total, 1);
            ExtractProgressBar.Value = 0;
            ProgressTextBlock.Text = $"진행 0 / {total}";
        }
    }

    private void UpdateExtractionProgress(int current, int total, string fileName)
    {
        ExtractProgressBar.Value = current;
        ProgressTextBlock.Text = $"진행 {current} / {total} - {fileName}";
    }

    private void ShowExportResultWindow(IReadOnlyCollection<string> createdFileNames)
    {
        string listText = createdFileNames.Count == 0
            ? "(생성된 파일 없음)"
            : string.Join(Environment.NewLine, createdFileNames
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .Select((name, idx) => $"{idx + 1}. {name}"));

        var window = new Window
        {
            Owner = this,
            Title = "JsonExporter 완료",
            Width = 560,
            Height = 420,
            MinWidth = 500,
            MinHeight = 320,
            WindowStartupLocation = WindowStartupLocation.CenterOwner
        };

        var root = new Grid { Margin = new Thickness(12) };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var title = new TextBlock
        {
            Text = $"변환 완료 파일 목록 ({createdFileNames.Count}개)",
            Margin = new Thickness(0, 0, 0, 8)
        };
        Grid.SetRow(title, 0);
        root.Children.Add(title);

        var listTextBox = new TextBox
        {
            IsReadOnly = true,
            Text = listText,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
            TextWrapping = TextWrapping.NoWrap
        };
        Grid.SetRow(listTextBox, 1);
        root.Children.Add(listTextBox);

        var okButton = new Button
        {
            Content = "확인",
            Width = 90,
            Height = 28,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 8, 0, 0)
        };
        okButton.Click += (_, _) => window.Close();
        Grid.SetRow(okButton, 2);
        root.Children.Add(okButton);

        window.Content = root;
        window.ShowDialog();
    }

    private void SaveExporterPaths()
    {
        if (_suppressPathSave)
        {
            return;
        }

        AppSettings settings = _settingsStore.Load();
        settings.JsonExporterExcelFolder = _excelFolder;
        settings.JsonExporterJsonFolder = _jsonFolder;
        _settingsStore.Save(settings);
    }

    private void JsonExtractNameItem_PropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (!string.Equals(e.PropertyName, nameof(JsonExtractNameItem.IsChecked), StringComparison.Ordinal))
        {
            return;
        }

        SaveCheckedNames();
    }

    private void SaveCheckedNames()
    {
        if (_suppressCheckedSave)
        {
            return;
        }

        AppSettings settings = _settingsStore.Load();
        settings.JsonExporterCheckedNames = _items
            .Where(x => x.IsChecked)
            .Select(x => x.Name)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
            .ToList();
        _settingsStore.Save(settings);
    }

    private static List<Dictionary<string, string>> BuildGroupRows(List<JsonExtractCandidate> files)
    {
        var allRows = new List<Dictionary<string, string>>();

        foreach (JsonExtractCandidate file in files
                     .OrderBy(x => x.FileStem, StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                using var workbook = new XLWorkbook(file.FilePath);
                IXLWorksheet sheet = workbook.Worksheet(1);
                IXLRange? used = sheet.RangeUsed();
                if (used is null)
                {
                    continue;
                }

                int firstRow = used.RangeAddress.FirstAddress.RowNumber;
                int lastRow = used.RangeAddress.LastAddress.RowNumber;
                int firstCol = used.RangeAddress.FirstAddress.ColumnNumber;
                int lastCol = used.RangeAddress.LastAddress.ColumnNumber;

                var exportColumns = new List<(int ColumnNumber, string HeaderName)>();
                var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int col = firstCol; col <= lastCol; col++)
                {
                    string rawHeader = sheet.Cell(firstRow, col).GetString().Trim();
                    if (rawHeader.StartsWith("#", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    string normalizedHeader = NormalizeHeader(rawHeader);
                    string header = string.IsNullOrWhiteSpace(normalizedHeader)
                        ? XLHelper.GetColumnLetterFromNumber(col)
                        : normalizedHeader;

                    string uniqueHeader = header;
                    int suffix = 2;
                    while (usedNames.Contains(uniqueHeader))
                    {
                        uniqueHeader = $"{header}_{suffix++}";
                    }

                    usedNames.Add(uniqueHeader);
                    exportColumns.Add((col, uniqueHeader));
                }

                for (int row = firstRow + 1; row <= lastRow; row++)
                {
                    var rowData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    bool hasAnyValue = false;

                    foreach ((int col, string headerName) in exportColumns)
                    {
                        string value = sheet.Cell(row, col).GetFormattedString();
                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            hasAnyValue = true;
                        }

                        rowData[headerName] = value;
                    }

                    if (hasAnyValue)
                    {
                        allRows.Add(rowData);
                    }
                }
            }
            catch
            {
                // 단일 파일 오류는 전체 추출을 중단하지 않기 위해 무시
            }
        }

        return allRows;
    }

    private static string GetItemCellValueByDisplayIndex(JsonExtractNameItem item, int displayIndex)
    {
        return displayIndex switch
        {
            0 => item.IsChecked ? "True" : "False",
            1 => item.Name,
            2 => item.GroupSummary,
            3 => item.FileCount.ToString(),
            4 => item.OutputJsonCount.ToString(),
            _ => string.Empty
        };
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

    private static string NormalizeHeader(string rawHeader)
    {
        string trimmed = rawHeader.Trim();
        int dollarIndex = trimmed.IndexOf('$');
        if (dollarIndex <= 0)
        {
            return trimmed;
        }

        return trimmed.Substring(0, dollarIndex).Trim();
    }
}

public sealed class JsonExtractNameItem : INotifyPropertyChanged
{
    private bool _isChecked;

    public bool IsChecked
    {
        get => _isChecked;
        set
        {
            if (_isChecked == value)
            {
                return;
            }

            _isChecked = value;
            OnPropertyChanged(nameof(IsChecked));
        }
    }

    public string Name { get; set; } = string.Empty;
    public string GroupSummary { get; set; } = string.Empty;
    public List<JsonExtractCandidate> Candidates { get; set; } = [];

    public int FileCount => Candidates.Count;
    public int OutputJsonCount => Candidates.Select(x => x.GroupKey).Distinct(StringComparer.OrdinalIgnoreCase).Count();

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged(string propertyName)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}

public sealed record JsonExtractCandidate(string FilePath, string FileStem, string Name, string GroupKey);
