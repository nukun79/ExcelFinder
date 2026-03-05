using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;
using System.Windows;
using Microsoft.Win32;

namespace ExcelFinder;

public partial class JsonDiffWindow : Window
{
    private readonly ObservableCollection<JsonDiffItem> _items = [];

    public JsonDiffWindow()
    {
        InitializeComponent();
        DiffDataGrid.ItemsSource = _items;
    }

    private void BrowseSourceButton_Click(object sender, RoutedEventArgs e)
    {
        string? path = PickJsonFile(SourcePathTextBox.Text.Trim(), "Source JSON 선택");
        if (path is not null)
        {
            SourcePathTextBox.Text = path;
        }
    }

    private void BrowseTargetButton_Click(object sender, RoutedEventArgs e)
    {
        string? path = PickJsonFile(TargetPathTextBox.Text.Trim(), "Target JSON 선택");
        if (path is not null)
        {
            TargetPathTextBox.Text = path;
        }
    }

    private void BrowseSourceDirButton_Click(object sender, RoutedEventArgs e)
    {
        string? path = PickFolder(SourceDirPathTextBox.Text.Trim(), "Source 폴더 선택");
        if (path is not null)
        {
            SourceDirPathTextBox.Text = path;
        }
    }

    private void BrowseTargetDirButton_Click(object sender, RoutedEventArgs e)
    {
        string? path = PickFolder(TargetDirPathTextBox.Text.Trim(), "Target 폴더 선택");
        if (path is not null)
        {
            TargetDirPathTextBox.Text = path;
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

    private void DiffButton_Click(object sender, RoutedEventArgs e)
    {
        _items.Clear();

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

            foreach (JsonDiffItem item in diffs)
            {
                _items.Add(item);
            }

            StatusTextBlock.Text = diffs.Count > 0
                ? $"차이점 {diffs.Count}건"
                : "차이점이 없습니다.";
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"Diff 실패: {ex.Message}";
        }
    }

    private void DiffFolderButton_Click(object sender, RoutedEventArgs e)
    {
        _items.Clear();

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
                foreach (JsonDiffItem item in diffs)
                {
                    _items.Add(item);
                }
            }

            StatusTextBlock.Text = _items.Count > 0
                ? $"차이점 {_items.Count}건 / 차이 파일 {diffFileCount}개 (비교 파일 {commonFiles.Count}개)"
                : $"차이점이 없습니다. (비교 파일 {commonFiles.Count}개)";
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"폴더 비교 실패: {ex.Message}";
        }
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
                foreach ((string field, string value) in targetRow!)
                {
                    result.Add(new JsonDiffItem
                    {
                        JsonFile = jsonFile,
                        DiffType = "추가",
                        Key = key,
                        FieldName = field,
                        SourceValue = string.Empty,
                        TargetValue = value
                    });
                }

                continue;
            }

            if (hasSource && !hasTarget)
            {
                foreach ((string field, string value) in sourceRow!)
                {
                    result.Add(new JsonDiffItem
                    {
                        JsonFile = jsonFile,
                        DiffType = "삭제",
                        Key = key,
                        FieldName = field,
                        SourceValue = value,
                        TargetValue = string.Empty
                    });
                }

                continue;
            }

            IEnumerable<string> fields = sourceRow!.Keys.Union(targetRow!.Keys, StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase);

            foreach (string field in fields)
            {
                string sourceValue = sourceRow.TryGetValue(field, out string? s) ? s : string.Empty;
                string targetValue = targetRow.TryGetValue(field, out string? t) ? t : string.Empty;

                if (string.Equals(sourceValue, targetValue, StringComparison.Ordinal))
                {
                    continue;
                }

                result.Add(new JsonDiffItem
                {
                    JsonFile = jsonFile,
                    DiffType = "변경",
                    Key = key,
                    FieldName = field,
                    SourceValue = sourceValue,
                    TargetValue = targetValue
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
