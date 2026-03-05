using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using ClosedXML.Excel;

namespace ExcelFinder;

public partial class HistoryDiffWindow : Window
{
    private readonly string _depotPath;
    private readonly int _rightRevision;
    private readonly ObservableCollection<ExcelDiffItem> _diffs = [];
    private readonly List<ExcelDiffItem> _allDiffs = [];
    private string _tempDir = string.Empty;
    private string _rightRevisionPath = string.Empty;

    public HistoryDiffWindow(string depotPath, int rightRevision)
    {
        InitializeComponent();

        _depotPath = depotPath;
        _rightRevision = rightRevision;
        DiffDataGrid.ItemsSource = _diffs;
        HeaderTextBlock.Text = $"History Diff: {_depotPath}#{_rightRevision - 1} -> #{_rightRevision}";

        Loaded += HistoryDiffWindow_Loaded;
        Closed += HistoryDiffWindow_Closed;
    }

    private async void HistoryDiffWindow_Loaded(object sender, RoutedEventArgs e)
    {
        await LoadDiffAsync();
    }

    private async Task LoadDiffAsync()
    {
        StatusTextBlock.Text = "Diff 조회 중...";

        CleanupTempFiles();
        _tempDir = Path.Combine(Path.GetTempPath(), "ExcelFinder", "HistoryDiff", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDir);

        string leftPath = Path.Combine(_tempDir, "left.xlsx");
        _rightRevisionPath = Path.Combine(_tempDir, "right.xlsx");

        try
        {
            string leftSpec = $"{_depotPath}#{_rightRevision - 1}";
            string rightSpec = $"{_depotPath}#{_rightRevision}";

            (bool leftOk, string leftMsg) = await Task.Run(() => PerforceHelper.ExportDepotRevisionToFile(leftSpec, leftPath));
            if (!leftOk)
            {
                StatusTextBlock.Text = $"이전 리비전 가져오기 실패: {leftMsg}";
                return;
            }

            (bool rightOk, string rightMsg) = await Task.Run(() => PerforceHelper.ExportDepotRevisionToFile(rightSpec, _rightRevisionPath));
            if (!rightOk)
            {
                StatusTextBlock.Text = $"선택 리비전 가져오기 실패: {rightMsg}";
                return;
            }

            List<ExcelDiffItem> diffs = await Task.Run(() => BuildDiffs(leftPath, _rightRevisionPath));
            _allDiffs.Clear();
            _allDiffs.AddRange(diffs);
            ApplyFindFilter();
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"Diff 오류: {ex.Message}";
        }
    }

    private void HistoryDiffWindow_Closed(object? sender, EventArgs e)
    {
        CleanupTempFiles();
    }

    private void CleanupTempFiles()
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(_tempDir) && Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, true);
            }
        }
        catch
        {
            // 임시 폴더 정리 실패는 무시
        }

        _tempDir = string.Empty;
        _rightRevisionPath = string.Empty;
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
            int? uidCol = FindUidColumnIndex(leftSheet, rightSheet, maxCol);

            for (int row = 1; row <= maxRow; row++)
            {
                bool leftRowHasValue = RowHasAnyValue(leftSheet, row, maxCol);
                bool rightRowHasValue = RowHasAnyValue(rightSheet, row, maxCol);
                string uidValue = GetUidValue(leftSheet, rightSheet, row, uidCol);

                // 이전 리비전에는 없고 선택 리비전에만 새로 생긴 행은 한 줄로 묶어서 보여준다.
                if (!leftRowHasValue && rightRowHasValue)
                {
                    result.Add(new ExcelDiffItem
                    {
                        SheetName = sheetName,
                        LineNumber = row,
                        UidValue = uidValue,
                        ColumnName = "(신규 행)",
                        CellAddress = $"ROW {row} (NEW)",
                        LeftValue = "(없음)",
                        RightValue = BuildRowContent(rightSheet, row, maxCol),
                        ResultValue = BuildRowContent(rightSheet, row, maxCol)
                    });
                    continue;
                }

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
                        LineNumber = row,
                        UidValue = uidValue,
                        ColumnName = GetDisplayColumnName(leftSheet, rightSheet, col),
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

    private static bool RowHasAnyValue(IXLWorksheet? sheet, int row, int maxCol)
    {
        if (sheet is null || maxCol <= 0)
        {
            return false;
        }

        for (int col = 1; col <= maxCol; col++)
        {
            string value = sheet.Cell(row, col).GetFormattedString();
            if (!string.IsNullOrWhiteSpace(value))
            {
                return true;
            }
        }

        return false;
    }

    private static string BuildRowContent(IXLWorksheet? sheet, int row, int maxCol)
    {
        if (sheet is null || maxCol <= 0)
        {
            return string.Empty;
        }

        var parts = new List<string>();
        for (int col = 1; col <= maxCol; col++)
        {
            string value = sheet.Cell(row, col).GetFormattedString();
            if (string.IsNullOrWhiteSpace(value))
            {
                continue;
            }

            string colName = GetDisplayColumnName(null, sheet, col);
            parts.Add($"{colName}:{value}");
        }

        return string.Join(" | ", parts);
    }

    private static int GetLastUsedRow(IXLWorksheet? sheet)
    {
        return sheet?.RangeUsed()?.RangeAddress.LastAddress.RowNumber ?? 0;
    }

    private static int GetLastUsedCol(IXLWorksheet? sheet)
    {
        return sheet?.RangeUsed()?.RangeAddress.LastAddress.ColumnNumber ?? 0;
    }

    private static int? FindUidColumnIndex(IXLWorksheet? leftSheet, IXLWorksheet? rightSheet, int maxCol)
    {
        if (maxCol <= 0)
        {
            return null;
        }

        for (int col = 1; col <= maxCol; col++)
        {
            string rightHeader = rightSheet?.Cell(1, col).GetFormattedString().Trim() ?? string.Empty;
            if (string.Equals(rightHeader, "UID", StringComparison.OrdinalIgnoreCase))
            {
                return col;
            }

            string leftHeader = leftSheet?.Cell(1, col).GetFormattedString().Trim() ?? string.Empty;
            if (string.Equals(leftHeader, "UID", StringComparison.OrdinalIgnoreCase))
            {
                return col;
            }
        }

        return null;
    }

    private static string GetUidValue(IXLWorksheet? leftSheet, IXLWorksheet? rightSheet, int row, int? uidCol)
    {
        if (row <= 0 || !uidCol.HasValue)
        {
            return string.Empty;
        }

        string rightValue = rightSheet?.Cell(row, uidCol.Value).GetFormattedString().Trim() ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(rightValue))
        {
            return rightValue;
        }

        return leftSheet?.Cell(row, uidCol.Value).GetFormattedString().Trim() ?? string.Empty;
    }

    private static string GetDisplayColumnName(IXLWorksheet? leftSheet, IXLWorksheet? rightSheet, int col)
    {
        string rightHeader = rightSheet?.Cell(1, col).GetFormattedString().Trim() ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(rightHeader))
        {
            return rightHeader;
        }

        string leftHeader = leftSheet?.Cell(1, col).GetFormattedString().Trim() ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(leftHeader))
        {
            return leftHeader;
        }

        return XLHelper.GetColumnLetterFromNumber(col);
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
        _diffs.Clear();

        if (_allDiffs.Count == 0)
        {
            StatusTextBlock.Text = "차이점이 없습니다.";
            return;
        }

        string keyword = FindTextBox.Text.Trim();
        IEnumerable<ExcelDiffItem> filtered = string.IsNullOrWhiteSpace(keyword)
            ? _allDiffs
            : _allDiffs.Where(x => IsMatch(x, keyword));

        foreach (ExcelDiffItem diff in filtered)
        {
            _diffs.Add(diff);
        }

        StatusTextBlock.Text = string.IsNullOrWhiteSpace(keyword)
            ? $"차이점 {_diffs.Count}건"
            : $"검색 결과 {_diffs.Count}건 / 전체 {_allDiffs.Count}건";
    }

    private static bool IsMatch(ExcelDiffItem item, string keyword)
    {
        return Contains(item.SheetName, keyword)
               || item.LineNumber.ToString().Contains(keyword, StringComparison.OrdinalIgnoreCase)
               || Contains(item.ColumnName, keyword)
               || Contains(item.LeftValue, keyword)
               || Contains(item.RightValue, keyword)
               || Contains(item.CellAddress, keyword);
    }

    private static bool Contains(string? text, string keyword)
    {
        return (text ?? string.Empty).Contains(keyword, StringComparison.OrdinalIgnoreCase);
    }

    private void DiffDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (DiffDataGrid.SelectedItem is not ExcelDiffItem)
        {
            return;
        }

        OpenRightRevisionFile();
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

    private void OpenEditorMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (DiffDataGrid.SelectedItem is not ExcelDiffItem selected)
        {
            StatusTextBlock.Text = "선택된 Diff 항목이 없습니다.";
            return;
        }

        if (string.IsNullOrWhiteSpace(_rightRevisionPath) || !File.Exists(_rightRevisionPath))
        {
            StatusTextBlock.Text = "리비전 파일이 준비되지 않아 에디터를 열 수 없습니다.";
            return;
        }

        if (selected.LineNumber <= 0)
        {
            StatusTextBlock.Text = "유효한 라인 번호가 없어 에디터를 열 수 없습니다.";
            return;
        }

        List<string> highlightAddresses = CollectLineHighlightAddresses(selected);
        var editor = new ExcelEditorWindow(_rightRevisionPath, selected.SheetName, selected.LineNumber, highlightAddresses)
        {
            Owner = this
        };
        editor.Show();
    }

    private void OpenRightRevisionFile()
    {
        if (string.IsNullOrWhiteSpace(_rightRevisionPath) || !File.Exists(_rightRevisionPath))
        {
            StatusTextBlock.Text = "리비전 파일이 준비되지 않아 Excel 파일을 열 수 없습니다.";
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = _rightRevisionPath,
                UseShellExecute = true
            });
            StatusTextBlock.Text = $"열기: {Path.GetFileName(_rightRevisionPath)}";
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"Excel 파일 열기 실패: {ex.Message}";
        }
    }

    private List<string> CollectLineHighlightAddresses(ExcelDiffItem selected)
    {
        return _allDiffs
            .Where(x => string.Equals(x.SheetName, selected.SheetName, StringComparison.OrdinalIgnoreCase)
                        && x.LineNumber == selected.LineNumber
                        && IsCellAddress(x.CellAddress))
            .Select(x => x.CellAddress)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static bool IsCellAddress(string? value)
    {
        return Regex.IsMatch(value ?? string.Empty, @"^[A-Za-z]+\d+$");
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
