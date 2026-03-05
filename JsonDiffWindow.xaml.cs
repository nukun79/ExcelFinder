using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;

namespace ExcelFinder;

public partial class JsonDiffWindow : Window
{
    private readonly AppSettingsStore _settingsStore = AppSettingsStore.CreateDefault();
    private readonly ObservableCollection<JsonDiffItem> _items = [];
    private readonly List<JsonDiffItem> _allDiffs = [];
    private bool _suppressPathSave;
    private string _contextCellValue = string.Empty;

    public JsonDiffWindow()
    {
        InitializeComponent();
        DiffDataGrid.ItemsSource = _items;

        LoadSavedPaths();
    }

    private void LoadSavedPaths()
    {
        AppSettings settings = _settingsStore.Load();
        _suppressPathSave = true;
        SourcePathTextBox.Text = settings.JsonDiffSourceFilePath ?? string.Empty;
        TargetPathTextBox.Text = settings.JsonDiffTargetFilePath ?? string.Empty;
        SourceDirPathTextBox.Text = settings.JsonDiffSourceDirPath ?? string.Empty;
        TargetDirPathTextBox.Text = settings.JsonDiffTargetDirPath ?? string.Empty;
        _suppressPathSave = false;
    }

    private void PathTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        SavePaths();
    }

    private void SavePaths()
    {
        if (_suppressPathSave)
        {
            return;
        }

        AppSettings settings = _settingsStore.Load();
        settings.JsonDiffSourceFilePath = SourcePathTextBox.Text.Trim();
        settings.JsonDiffTargetFilePath = TargetPathTextBox.Text.Trim();
        settings.JsonDiffSourceDirPath = SourceDirPathTextBox.Text.Trim();
        settings.JsonDiffTargetDirPath = TargetDirPathTextBox.Text.Trim();
        _settingsStore.Save(settings);
    }

    private void BrowseSourceButton_Click(object sender, RoutedEventArgs e)
    {
        string? path = PickJsonFile(SourcePathTextBox.Text.Trim(), "Source JSON 선택");
        if (path is not null)
        {
            SourcePathTextBox.Text = path;
            SavePaths();
        }
    }

    private void BrowseTargetButton_Click(object sender, RoutedEventArgs e)
    {
        string? path = PickJsonFile(TargetPathTextBox.Text.Trim(), "Target JSON 선택");
        if (path is not null)
        {
            TargetPathTextBox.Text = path;
            SavePaths();
        }
    }

    private void BrowseSourceDirButton_Click(object sender, RoutedEventArgs e)
    {
        string? path = PickFolder(SourceDirPathTextBox.Text.Trim(), "Source 폴더 선택");
        if (path is not null)
        {
            SourceDirPathTextBox.Text = path;
            SavePaths();
        }
    }

    private void BrowseTargetDirButton_Click(object sender, RoutedEventArgs e)
    {
        string? path = PickFolder(TargetDirPathTextBox.Text.Trim(), "Target 폴더 선택");
        if (path is not null)
        {
            TargetDirPathTextBox.Text = path;
            SavePaths();
        }
    }

    private static string? PickJsonFile(string currentPath, string title)
    {
        string initialDir = System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(currentPath) ?? string.Empty)
            ? System.IO.Path.GetDirectoryName(currentPath) ?? string.Empty
            : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var dialog = new OpenFileDialog
        {
            Title = title,
            Filter = "JSON Files|*.json|All Files|*.*",
            CheckFileExists = true,
            InitialDirectory = initialDir
        };

        bool? result = dialog.ShowDialog();
        return result == true ? dialog.FileName : null;
    }

    private static string? PickFolder(string currentPath, string title)
    {
        string initialDir = Directory.Exists(currentPath)
            ? currentPath
            : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var dialog = new OpenFolderDialog
        {
            Title = title,
            InitialDirectory = initialDir
        };

        bool? result = dialog.ShowDialog();
        return result == true ? dialog.FolderName : null;
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void DiffDataGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        _contextCellValue = string.Empty;

        DependencyObject? source = e.OriginalSource as DependencyObject;
        if (source is null)
        {
            return;
        }

        DataGridCell? cell = FindVisualParent<DataGridCell>(source);
        DataGridRow? row = FindVisualParent<DataGridRow>(source);
        if (row?.Item is not JsonDiffItem item)
        {
            return;
        }

        row.IsSelected = true;
        DiffDataGrid.SelectedItem = row.Item;

        int displayIndex = cell?.Column?.DisplayIndex ?? -1;
        _contextCellValue = GetCellValueByDisplayIndex(item, displayIndex);
    }

    private void CopyCellValueMenuItem_Click(object sender, RoutedEventArgs e)
    {
        string value = _contextCellValue;
        if (string.IsNullOrEmpty(value) && DiffDataGrid.SelectedItem is JsonDiffItem selected)
        {
            value = selected.TargetValue;
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

    private void DiffButton_Click(object sender, RoutedEventArgs e)
    {
        _items.Clear();
        _allDiffs.Clear();

        string sourcePath = SourcePathTextBox.Text.Trim();
        string targetPath = TargetPathTextBox.Text.Trim();

        if (!System.IO.File.Exists(sourcePath))
        {
            StatusTextBlock.Text = "Source 파일 경로를 확인해 주세요.";
            return;
        }

        if (!System.IO.File.Exists(targetPath))
        {
            StatusTextBlock.Text = "Target 파일 경로를 확인해 주세요.";
            return;
        }

        try
        {
            List<Dictionary<string, string>> sourceRows = LoadRows(sourcePath);
            List<Dictionary<string, string>> targetRows = LoadRows(targetPath);
            string fileLabel = $"{Path.GetFileName(sourcePath)} ↔ {Path.GetFileName(targetPath)}";
            List<JsonDiffItem> diffs = BuildDiffs(sourceRows, targetRows, fileLabel);

            _allDiffs.AddRange(diffs);
            ApplyFindFilter();
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"Diff 실패: {ex.Message}";
        }
    }

    private void DiffFolderButton_Click(object sender, RoutedEventArgs e)
    {
        _items.Clear();
        _allDiffs.Clear();

        string sourceDir = SourceDirPathTextBox.Text.Trim();
        string targetDir = TargetDirPathTextBox.Text.Trim();
        if (!Directory.Exists(sourceDir))
        {
            StatusTextBlock.Text = "Source Path 폴더 경로를 확인해 주세요.";
            return;
        }

        if (!Directory.Exists(targetDir))
        {
            StatusTextBlock.Text = "Target Path 폴더 경로를 확인해 주세요.";
            return;
        }

        try
        {
            Dictionary<string, string> sourceFiles = EnumerateJsonFilesByRelativePath(sourceDir);
            Dictionary<string, string> targetFiles = EnumerateJsonFilesByRelativePath(targetDir);
            List<string> commonFiles = sourceFiles.Keys
                .Intersect(targetFiles.Keys, StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (commonFiles.Count == 0)
            {
                StatusTextBlock.Text = "두 폴더에 같은 JSON 파일이 없습니다.";
                return;
            }

            int diffFileCount = 0;
            foreach (string relativePath in commonFiles)
            {
                List<Dictionary<string, string>> sourceRows = LoadRows(sourceFiles[relativePath]);
                List<Dictionary<string, string>> targetRows = LoadRows(targetFiles[relativePath]);
                List<JsonDiffItem> diffs = BuildDiffs(sourceRows, targetRows, relativePath);
                if (diffs.Count == 0)
                {
                    continue;
                }

                diffFileCount++;
                _allDiffs.AddRange(diffs);
            }

            ApplyFindFilter();
            if (_allDiffs.Count == 0)
            {
                StatusTextBlock.Text = $"차이점이 없습니다. (비교 파일 {commonFiles.Count}개)";
                return;
            }

            StatusTextBlock.Text = $"차이점 {_allDiffs.Count}건 / 차이 파일 {diffFileCount}개 (비교 파일 {commonFiles.Count}개)";
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"폴더 비교 실패: {ex.Message}";
        }
    }

    private void FindButton_Click(object sender, RoutedEventArgs e)
    {
        ApplyFindFilter();
    }

    private void FindTextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter)
        {
            return;
        }

        ApplyFindFilter();
        e.Handled = true;
    }

    private void ApplyFindFilter()
    {
        _items.Clear();

        if (_allDiffs.Count == 0)
        {
            StatusTextBlock.Text = "차이점이 없습니다.";
            return;
        }

        string keyword = FindTextBox.Text.Trim();
        IEnumerable<JsonDiffItem> filtered = string.IsNullOrWhiteSpace(keyword)
            ? _allDiffs
            : _allDiffs.Where(x => IsMatch(x, keyword));

        foreach (JsonDiffItem item in filtered)
        {
            _items.Add(item);
        }

        StatusTextBlock.Text = string.IsNullOrWhiteSpace(keyword)
            ? $"차이점 {_allDiffs.Count}건"
            : $"검색 결과 {_items.Count}건 / 전체 {_allDiffs.Count}건";
    }

    private static bool IsMatch(JsonDiffItem item, string keyword)
    {
        return Contains(item.JsonFile, keyword)
               || Contains(item.DiffType, keyword)
               || Contains(item.Key, keyword)
               || Contains(item.FieldName, keyword)
               || Contains(item.SourceValue, keyword)
               || Contains(item.TargetValue, keyword);
    }

    private static bool Contains(string? text, string keyword)
    {
        return (text ?? string.Empty).Contains(keyword, StringComparison.OrdinalIgnoreCase);
    }

    private static List<Dictionary<string, string>> LoadRows(string path)
    {
        string json = System.IO.File.ReadAllText(path);
        using JsonDocument doc = JsonDocument.Parse(json);
        if (doc.RootElement.ValueKind != JsonValueKind.Array)
        {
            throw new InvalidOperationException("JSON 최상위 구조는 배열이어야 합니다.");
        }

        var rows = new List<Dictionary<string, string>>();
        foreach (JsonElement row in doc.RootElement.EnumerateArray())
        {
            if (row.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            var item = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (JsonProperty prop in row.EnumerateObject())
            {
                item[prop.Name] = ToDisplayString(prop.Value);
            }

            rows.Add(item);
        }

        return rows;
    }

    private static string ToDisplayString(JsonElement value)
    {
        return value.ValueKind switch
        {
            JsonValueKind.Null => string.Empty,
            JsonValueKind.String => value.GetString() ?? string.Empty,
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            _ => value.GetRawText()
        };
    }

    private static List<JsonDiffItem> BuildDiffs(
        List<Dictionary<string, string>> sourceRows,
        List<Dictionary<string, string>> targetRows,
        string jsonFile)
    {
        var result = new List<JsonDiffItem>();

        var sourceMap = BuildRowMap(sourceRows);
        var targetMap = BuildRowMap(targetRows);
        var allKeys = sourceMap.Keys.Union(targetMap.Keys, StringComparer.OrdinalIgnoreCase)
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase);

        foreach (string key in allKeys)
        {
            bool hasSource = sourceMap.TryGetValue(key, out Dictionary<string, string>? sourceRow);
            bool hasTarget = targetMap.TryGetValue(key, out Dictionary<string, string>? targetRow);

            if (!hasSource && hasTarget)
            {
                Dictionary<string, string> addedFields = targetRow!
                    .OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(x => x.Key, x => x.Value, StringComparer.OrdinalIgnoreCase);
                result.Add(new JsonDiffItem
                {
                    JsonFile = jsonFile,
                    DiffType = "추가",
                    Key = key,
                    FieldName = string.Join(" | ", addedFields.Keys),
                    SourceValue = string.Empty,
                    TargetValue = BuildSummaryText(addedFields)
                });
                continue;
            }

            if (hasSource && !hasTarget)
            {
                Dictionary<string, string> deletedFields = sourceRow!
                    .OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(x => x.Key, x => x.Value, StringComparer.OrdinalIgnoreCase);
                result.Add(new JsonDiffItem
                {
                    JsonFile = jsonFile,
                    DiffType = "삭제",
                    Key = key,
                    FieldName = string.Join(" | ", deletedFields.Keys),
                    SourceValue = BuildSummaryText(deletedFields),
                    TargetValue = string.Empty
                });
                continue;
            }

            IEnumerable<string> fields = sourceRow!.Keys.Union(targetRow!.Keys, StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase);

            var changedSource = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var changedTarget = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string field in fields)
            {
                string sourceValue = sourceRow.TryGetValue(field, out string? s) ? s : string.Empty;
                string targetValue = targetRow.TryGetValue(field, out string? t) ? t : string.Empty;

                if (string.Equals(sourceValue, targetValue, StringComparison.Ordinal))
                {
                    continue;
                }

                changedSource[field] = sourceValue;
                changedTarget[field] = targetValue;
            }

            if (changedSource.Count > 0)
            {
                result.Add(new JsonDiffItem
                {
                    JsonFile = jsonFile,
                    DiffType = "변경",
                    Key = key,
                    FieldName = string.Join(" | ", changedSource.Keys.OrderBy(x => x, StringComparer.OrdinalIgnoreCase)),
                    SourceValue = BuildSummaryText(changedSource),
                    TargetValue = BuildSummaryText(changedTarget)
                });
            }
        }

        return result;
    }

    private static Dictionary<string, Dictionary<string, string>> BuildRowMap(List<Dictionary<string, string>> rows)
    {
        var map = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < rows.Count; i++)
        {
            Dictionary<string, string> row = rows[i];
            string key = row.TryGetValue("UID", out string? uid) && !string.IsNullOrWhiteSpace(uid)
                ? $"UID:{uid}"
                : $"INDEX:{i + 1}";

            map[key] = row;
        }

        return map;
    }

    private static Dictionary<string, string> EnumerateJsonFilesByRelativePath(string rootPath)
    {
        return Directory.EnumerateFiles(rootPath, "*.json", SearchOption.AllDirectories)
            .ToDictionary(
                path => Path.GetRelativePath(rootPath, path).Replace('\\', '/'),
                path => path,
                StringComparer.OrdinalIgnoreCase);
    }

    private static string BuildSummaryText(Dictionary<string, string> fields)
    {
        return string.Join(" | ",
            fields.OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase)
                .Select(x => $"{x.Key}:{x.Value}"));
    }

    private static string GetCellValueByDisplayIndex(JsonDiffItem item, int displayIndex)
    {
        return displayIndex switch
        {
            0 => item.JsonFile,
            1 => item.DiffType,
            2 => item.Key,
            3 => item.FieldName,
            4 => item.SourceValue,
            5 => item.TargetValue,
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
}

public sealed class JsonDiffItem
{
    public string JsonFile { get; set; } = string.Empty;
    public string DiffType { get; set; } = string.Empty;
    public string Key { get; set; } = string.Empty;
    public string FieldName { get; set; } = string.Empty;
    public string SourceValue { get; set; } = string.Empty;
    public string TargetValue { get; set; } = string.Empty;
}
