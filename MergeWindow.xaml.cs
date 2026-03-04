using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace ExcelFinder;

public partial class MergeWindow : Window
{
    private readonly ObservableCollection<ExcelDiffItem> _diffs = [];

    public MergeWindow()
    {
        InitializeComponent();
        DiffDataGrid.ItemsSource = _diffs;
    }

    private void BrowseLeftButton_Click(object sender, RoutedEventArgs e)
    {
        string? path = PickExcelFile(LeftFileTextBox.Text);
        if (path is not null)
        {
            LeftFileTextBox.Text = path;
        }
    }

    private void BrowseRightButton_Click(object sender, RoutedEventArgs e)
    {
        string? path = PickExcelFile(RightFileTextBox.Text);
        if (path is not null)
        {
            RightFileTextBox.Text = path;
        }
    }

    private string? PickExcelFile(string currentPath)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Excel 파일 선택",
            Filter = "Excel Files|*.xlsx;*.xlsm",
            CheckFileExists = true,
            InitialDirectory = Directory.Exists(Path.GetDirectoryName(currentPath) ?? string.Empty)
                ? Path.GetDirectoryName(currentPath)
                : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        };

        bool? result = dialog.ShowDialog(this);
        return result == true ? dialog.FileName : null;
    }

    private void CompareButton_Click(object sender, RoutedEventArgs e)
    {
        _diffs.Clear();

        if (!ValidateInputFiles(out string leftPath, out string rightPath))
        {
            return;
        }

        try
        {
            List<ExcelDiffItem> diffs = BuildDiffs(leftPath, rightPath);
            foreach (ExcelDiffItem diff in diffs)
            {
                _diffs.Add(diff);
            }

            StatusTextBlock.Text = diffs.Count > 0
                ? $"차이점 {diffs.Count}건"
                : "차이점이 없습니다.";
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"비교 실패: {ex.Message}";
        }
    }

    private void MergeButton_Click(object sender, RoutedEventArgs e)
    {
        if (!ValidateInputFiles(out string leftPath, out string rightPath))
        {
            return;
        }

        try
        {
            List<ExcelDiffItem> diffs = _diffs.Count > 0
                ? _diffs.ToList()
                : BuildDiffs(leftPath, rightPath);
            if (diffs.Count == 0)
            {
                StatusTextBlock.Text = "차이점이 없어 머지할 내용이 없습니다.";
                return;
            }

            string mergedPath = CreateMergedFile(leftPath, rightPath, diffs);
            StatusTextBlock.Text = $"머지 완료: {mergedPath}";

            _diffs.Clear();
            foreach (ExcelDiffItem diff in diffs)
            {
                _diffs.Add(diff);
            }
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"머지 실패: {ex.Message}";
        }
    }

    private void DiffDataGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        DependencyObject? source = e.OriginalSource as DependencyObject;
        if (source is null)
        {
            return;
        }

        DataGridRow? row = FindVisualParent<DataGridRow>(source);
        if (row?.Item is ExcelDiffItem)
        {
            row.IsSelected = true;
            DiffDataGrid.SelectedItem = row.Item;
        }
    }

    private void OpenLeftEditorMenuItem_Click(object sender, RoutedEventArgs e)
    {
        OpenEditorForSelectedDiff(LeftFileTextBox.Text.Trim(), "기준");
    }

    private void OpenRightEditorMenuItem_Click(object sender, RoutedEventArgs e)
    {
        OpenEditorForSelectedDiff(RightFileTextBox.Text.Trim(), "비교");
    }

    private void OpenEditorForSelectedDiff(string targetPath, string label)
    {
        if (DiffDataGrid.SelectedItem is not ExcelDiffItem diff)
        {
            StatusTextBlock.Text = "선택된 차이점 항목이 없습니다.";
            return;
        }

        if (!File.Exists(targetPath))
        {
            StatusTextBlock.Text = $"{label} Excel 파일 경로를 확인해 주세요.";
            return;
        }

        int rowNumber = ParseRowNumber(diff.CellAddress);
        if (rowNumber <= 0)
        {
            StatusTextBlock.Text = $"셀 주소에서 행 번호를 읽을 수 없습니다: {diff.CellAddress}";
            return;
        }

        var editor = new ExcelEditorWindow(targetPath, diff.SheetName, rowNumber, [diff.CellAddress])
        {
            Owner = this
        };
        editor.Show();
    }

    private bool ValidateInputFiles(out string leftPath, out string rightPath)
    {
        leftPath = LeftFileTextBox.Text.Trim();
        rightPath = RightFileTextBox.Text.Trim();

        if (!File.Exists(leftPath))
        {
            StatusTextBlock.Text = "기준 Excel 파일 경로를 확인해 주세요.";
            return false;
        }

        if (!File.Exists(rightPath))
        {
            StatusTextBlock.Text = "비교 Excel 파일 경로를 확인해 주세요.";
            return false;
        }

        return true;
    }

    private static List<ExcelDiffItem> BuildDiffs(string leftPath, string rightPath)
    {
        var result = new List<ExcelDiffItem>();

        using var leftWb = new XLWorkbook(leftPath);
        using var rightWb = new XLWorkbook(rightPath);

        List<string> sheetNames = leftWb.Worksheets.Select(w => w.Name)
            .Union(rightWb.Worksheets.Select(w => w.Name), StringComparer.OrdinalIgnoreCase)
            .ToList();

        foreach (string sheetName in sheetNames)
        {
            IXLWorksheet? leftSheet = leftWb.Worksheets.FirstOrDefault(w => string.Equals(w.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            IXLWorksheet? rightSheet = rightWb.Worksheets.FirstOrDefault(w => string.Equals(w.Name, sheetName, StringComparison.OrdinalIgnoreCase));

            int maxRow = Math.Max(GetLastUsedRow(leftSheet), GetLastUsedRow(rightSheet));
            int maxCol = Math.Max(GetLastUsedCol(leftSheet), GetLastUsedCol(rightSheet));

            for (int row = 1; row <= maxRow; row++)
            {
                for (int col = 1; col <= maxCol; col++)
                {
                    string leftValue = leftSheet?.Cell(row, col).GetFormattedString() ?? string.Empty;
                    string rightValue = rightSheet?.Cell(row, col).GetFormattedString() ?? string.Empty;

                    if (string.Equals(leftValue, rightValue, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    result.Add(new ExcelDiffItem
                    {
                        SheetName = sheetName,
                        CellAddress = XLHelper.GetColumnLetterFromNumber(col) + row,
                        LeftValue = leftValue,
                        RightValue = rightValue,
                        ResultValue = rightValue
                    });
                }
            }
        }

        return result;
    }

    private static int GetLastUsedRow(IXLWorksheet? sheet)
    {
        return sheet?.RangeUsed()?.RangeAddress.LastAddress.RowNumber ?? 0;
    }

    private static int GetLastUsedCol(IXLWorksheet? sheet)
    {
        return sheet?.RangeUsed()?.RangeAddress.LastAddress.ColumnNumber ?? 0;
    }

    private static int ParseRowNumber(string cellAddress)
    {
        Match m = Regex.Match(cellAddress ?? string.Empty, @"(\d+)$");
        return m.Success && int.TryParse(m.Groups[1].Value, out int row) ? row : -1;
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

    private static string CreateMergedFile(string leftPath, string rightPath, IReadOnlyList<ExcelDiffItem> diffs)
    {
        string dir = Path.GetDirectoryName(leftPath) ?? Environment.CurrentDirectory;
        string fileName = Path.GetFileNameWithoutExtension(leftPath);
        string ext = Path.GetExtension(leftPath);
        string mergedPath = Path.Combine(dir, $"{fileName}_merged{ext}");

        int index = 2;
        while (File.Exists(mergedPath))
        {
            mergedPath = Path.Combine(dir, $"{fileName}_merged_{index}{ext}");
            index++;
        }

        File.Copy(leftPath, mergedPath, false);

        using var mergedWb = new XLWorkbook(mergedPath);
        using var rightWb = new XLWorkbook(rightPath);
        foreach (ExcelDiffItem diff in diffs)
        {
            IXLWorksheet mergedSheet = mergedWb.Worksheets.FirstOrDefault(w => string.Equals(w.Name, diff.SheetName, StringComparison.OrdinalIgnoreCase))
                                     ?? mergedWb.AddWorksheet(diff.SheetName);
            IXLCell mergedCell = mergedSheet.Cell(diff.CellAddress);

            // 결과값이 비교값과 동일하면 비교 셀 자체를 복사해 타입/서식/수식까지 그대로 반영한다.
            if (string.Equals(diff.ResultValue, diff.RightValue, StringComparison.Ordinal))
            {
                IXLWorksheet? rightSheet = rightWb.Worksheets.FirstOrDefault(w => string.Equals(w.Name, diff.SheetName, StringComparison.OrdinalIgnoreCase));
                if (rightSheet is not null)
                {
                    rightSheet.Cell(diff.CellAddress).CopyTo(mergedCell);
                    continue;
                }
            }

            SetTypedValue(mergedCell, diff.ResultValue ?? string.Empty);
        }

        mergedWb.Save();
        return mergedPath;
    }

    private static void SetTypedValue(IXLCell cell, string raw)
    {
        string value = raw.Trim();
        if (string.IsNullOrEmpty(value))
        {
            cell.Clear(XLClearOptions.Contents);
            return;
        }

        if (bool.TryParse(value, out bool boolValue))
        {
            cell.Value = boolValue;
            return;
        }

        if (long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long longValue))
        {
            cell.Value = longValue;
            return;
        }

        if (double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out double doubleValue))
        {
            cell.Value = doubleValue;
            return;
        }

        if (double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out double localDouble))
        {
            cell.Value = localDouble;
            return;
        }

        if (DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.None, out DateTime dt))
        {
            cell.Value = dt;
            return;
        }

        cell.Value = value;
    }
}

public sealed class ExcelDiffItem
{
    public string SheetName { get; set; } = string.Empty;
    public int LineNumber { get; set; }
    public string ColumnName { get; set; } = string.Empty;
    public string CellAddress { get; set; } = string.Empty;
    public string LeftValue { get; set; } = string.Empty;
    public string RightValue { get; set; } = string.Empty;
    public string ResultValue { get; set; } = string.Empty;
}
