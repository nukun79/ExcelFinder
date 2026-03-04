using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using ClosedXML.Excel;

namespace ExcelFinder;

public partial class HistoryDiffWindow : Window
{
    private readonly string _depotPath;
    private readonly int _rightRevision;
    private readonly ObservableCollection<ExcelDiffItem> _diffs = [];

    public HistoryDiffWindow(string depotPath, int rightRevision)
    {
        InitializeComponent();

        _depotPath = depotPath;
        _rightRevision = rightRevision;
        DiffDataGrid.ItemsSource = _diffs;
        HeaderTextBlock.Text = $"History Diff: {_depotPath}#{_rightRevision - 1} -> #{_rightRevision}";

        Loaded += HistoryDiffWindow_Loaded;
    }

    private async void HistoryDiffWindow_Loaded(object sender, RoutedEventArgs e)
    {
        await LoadDiffAsync();
    }

    private async Task LoadDiffAsync()
    {
        StatusTextBlock.Text = "Diff 조회 중...";

        string tempDir = Path.Combine(Path.GetTempPath(), "ExcelFinder", "HistoryDiff", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        string leftPath = Path.Combine(tempDir, "left.xlsx");
        string rightPath = Path.Combine(tempDir, "right.xlsx");

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

            (bool rightOk, string rightMsg) = await Task.Run(() => PerforceHelper.ExportDepotRevisionToFile(rightSpec, rightPath));
            if (!rightOk)
            {
                StatusTextBlock.Text = $"선택 리비전 가져오기 실패: {rightMsg}";
                return;
            }

            List<ExcelDiffItem> diffs = await Task.Run(() => BuildDiffs(leftPath, rightPath));
            _diffs.Clear();
            foreach (ExcelDiffItem diff in diffs)
            {
                _diffs.Add(diff);
            }

            StatusTextBlock.Text = diffs.Count > 0 ? $"차이점 {diffs.Count}건" : "차이점이 없습니다.";
        }
        catch (Exception ex)
        {
            StatusTextBlock.Text = $"Diff 오류: {ex.Message}";
        }
        finally
        {
            try
            {
                if (Directory.Exists(tempDir))
                {
                    Directory.Delete(tempDir, true);
                }
            }
            catch
            {
                // 임시 폴더 정리 실패는 무시
            }
        }
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
                bool leftRowHasValue = RowHasAnyValue(leftSheet, row, maxCol);
                bool rightRowHasValue = RowHasAnyValue(rightSheet, row, maxCol);

                // 이전 리비전에는 없고 선택 리비전에만 새로 생긴 행은 한 줄로 묶어서 보여준다.
                if (!leftRowHasValue && rightRowHasValue)
                {
                    result.Add(new ExcelDiffItem
                    {
                        SheetName = sheetName,
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

            string colName = XLHelper.GetColumnLetterFromNumber(col);
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
}
