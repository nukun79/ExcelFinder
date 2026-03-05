using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using System.Windows.Threading;
using ClosedXML.Excel;

namespace ExcelFinder;

public partial class ExcelEditorWindow : Window
{
    private readonly string _filePath;
    private readonly string _sheetName;
    private readonly int _targetRowNumber;
    private readonly AppSettingsStore _settingsStore = AppSettingsStore.CreateDefault();
    private readonly HashSet<int> _highlightExcelColumnNumbers = [];

    private XLWorkbook? _workbook;
    private IXLWorksheet? _worksheet;
    private DataTable? _table;
    private readonly List<int> _sheetColumnNumbers = [];
    private readonly List<int> _sheetRowNumbers = [];
    private readonly Dictionary<string, string> _originalSnapshot = new(StringComparer.Ordinal);
    private readonly HashSet<int> _highlightTableColumnIndexes = [];
    private int _targetTableRowIndex = -1;

    public ExcelEditorWindow(string filePath, string sheetName, int targetRowNumber, IEnumerable<string>? highlightCellAddresses = null)
    {
        InitializeComponent();

        _filePath = filePath;
        _sheetName = sheetName;
        _targetRowNumber = targetRowNumber;

        if (highlightCellAddresses is not null)
        {
            foreach (string address in highlightCellAddresses)
            {
                int col = GetColumnIndexFromCellAddress(address);
                if (col > 0)
                {
                    _highlightExcelColumnNumbers.Add(col);
                }
            }
        }

        Title = $"Excel Editor - {Path.GetFileName(_filePath)}";
        FileInfoTextBlock.Text = $"파일: {_filePath}\n시트: {_sheetName} / 검색 행: {_targetRowNumber}";
        CheckinDescriptionTextBox.Text = $"ExcelFinder edit: {Path.GetFileName(_filePath)}";
        LoadSavedCheckinPrefix();

        Loaded += ExcelEditorWindow_Loaded;
        Closed += ExcelEditorWindow_Closed;
    }

    private void ExcelEditorWindow_Loaded(object sender, RoutedEventArgs e)
    {
        LoadSheetData();
        RefreshCheckoutState();
    }

    private void ExcelEditorWindow_Closed(object? sender, EventArgs e)
    {
        _workbook?.Dispose();
    }

    private void CheckoutButton_Click(object sender, RoutedEventArgs e)
    {
        EnsureCheckout();
    }

    private void EnsureCheckout()
    {
        (bool success, string message) = PerforceHelper.Checkout(_filePath);
        if (success)
        {
            EditorStatusTextBlock.Text = "Perforce 체크아웃 완료";
            RefreshCheckoutState();
            return;
        }

        EditorStatusTextBlock.Text = $"체크아웃 실패: {message}";
        (bool infoOk, string _, PerforceClientSettings? settings) = PerforceHelper.GetClientSettings();
        var configWindow = new PerforceConfigWindow(infoOk ? settings : null)
        {
            Owner = this
        };

        bool? result = configWindow.ShowDialog();
        if (result == true)
        {
            (bool retrySuccess, string retryMessage) = PerforceHelper.Checkout(_filePath);
            EditorStatusTextBlock.Text = retrySuccess
                ? "Perforce 체크아웃 완료"
                : $"재시도 실패: {retryMessage}";
            RefreshCheckoutState();
        }
    }

    private void LoadSheetData()
    {
        _workbook?.Dispose();
        _workbook = new XLWorkbook(_filePath);

        _worksheet = _workbook.Worksheets.FirstOrDefault(w => string.Equals(w.Name, _sheetName, StringComparison.OrdinalIgnoreCase))
                     ?? _workbook.Worksheet(1);

        IXLRange? used = _worksheet.RangeUsed();
        if (used is null)
        {
            _table = new DataTable();
            SheetDataGrid.ItemsSource = _table.DefaultView;
            EditorStatusTextBlock.Text = "시트가 비어 있습니다.";
            return;
        }

        int firstRow = used.RangeAddress.FirstAddress.RowNumber;
        int lastRow = used.RangeAddress.LastAddress.RowNumber;
        int firstCol = used.RangeAddress.FirstAddress.ColumnNumber;
        int lastCol = used.RangeAddress.LastAddress.ColumnNumber;

        int headerRow = firstRow;
        _table = new DataTable();
        _sheetColumnNumbers.Clear();
        _sheetRowNumbers.Clear();
        _highlightTableColumnIndexes.Clear();

        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int col = firstCol; col <= lastCol; col++)
        {
            string rawName = _worksheet.Cell(headerRow, col).GetString().Trim();
            string columnName = string.IsNullOrWhiteSpace(rawName)
                ? XLHelper.GetColumnLetterFromNumber(col)
                : rawName;

            string uniqueName = columnName;
            int suffix = 2;
            while (usedNames.Contains(uniqueName))
            {
                uniqueName = $"{columnName}_{suffix++}";
            }

            usedNames.Add(uniqueName);
            _table.Columns.Add(uniqueName, typeof(string));
            _sheetColumnNumbers.Add(col);
        }

        _targetTableRowIndex = -1;

        for (int row = headerRow + 1; row <= lastRow; row++)
        {
            DataRow dataRow = _table.NewRow();
            for (int c = 0; c < _sheetColumnNumbers.Count; c++)
            {
                int excelCol = _sheetColumnNumbers[c];
                dataRow[c] = _worksheet.Cell(row, excelCol).GetFormattedString();
            }

            _table.Rows.Add(dataRow);
            _sheetRowNumbers.Add(row);

            if (row == _targetRowNumber)
            {
                _targetTableRowIndex = _table.Rows.Count - 1;
            }
        }

        SheetDataGrid.ItemsSource = _table.DefaultView;
        CaptureOriginalSnapshot();

        for (int i = 0; i < _sheetColumnNumbers.Count; i++)
        {
            if (_highlightExcelColumnNumbers.Contains(_sheetColumnNumbers[i]))
            {
                _highlightTableColumnIndexes.Add(i);
            }
        }

        if (_targetTableRowIndex >= 0)
        {
            SheetDataGrid.SelectedIndex = _targetTableRowIndex;
            SheetDataGrid.ScrollIntoView(SheetDataGrid.Items[_targetTableRowIndex]);
        }

        EditorStatusTextBlock.Text = $"로드 완료: {_worksheet.Name} ({_table.Rows.Count}행)";
    }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        if (_workbook is null || _worksheet is null || _table is null)
        {
            EditorStatusTextBlock.Text = "저장할 데이터가 없습니다.";
            return;
        }

        for (int r = 0; r < _table.Rows.Count; r++)
        {
            int excelRow = _sheetRowNumbers[r];
            for (int c = 0; c < _sheetColumnNumbers.Count; c++)
            {
                int excelCol = _sheetColumnNumbers[c];
                string value = _table.Rows[r][c]?.ToString() ?? string.Empty;
                _worksheet.Cell(excelRow, excelCol).Value = value;
            }
        }

        _workbook.Save();
        EditorStatusTextBlock.Text = "저장 완료";
    }

    private void DiffButton_Click(object sender, RoutedEventArgs e)
    {
        List<CellDiffItem> diffs = GetCurrentDiffItems();
        if (diffs.Count == 0)
        {
            EditorStatusTextBlock.Text = "변경된 내용이 없습니다.";
            return;
        }

        var diffWindow = new Window
        {
            Owner = this,
            Title = $"Diff - {Path.GetFileName(_filePath)} ({diffs.Count}건)",
            Width = 1000,
            Height = 640,
            MinWidth = 760,
            MinHeight = 420
        };

        var grid = new System.Windows.Controls.Grid { Margin = new Thickness(12) };
        grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = GridLength.Auto });
        grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

        var infoText = new System.Windows.Controls.TextBlock
        {
            Text = $"변경 건수: {diffs.Count}",
            Margin = new Thickness(0, 0, 0, 8)
        };
        System.Windows.Controls.Grid.SetRow(infoText, 0);
        grid.Children.Add(infoText);

        var diffGrid = new System.Windows.Controls.DataGrid
        {
            AutoGenerateColumns = false,
            IsReadOnly = true,
            CanUserAddRows = false,
            CanUserDeleteRows = false,
            CanUserResizeRows = false,
            ItemsSource = diffs
        };
        diffGrid.Columns.Add(new System.Windows.Controls.DataGridTextColumn { Header = "Excel 행", Binding = new System.Windows.Data.Binding(nameof(CellDiffItem.ExcelRow)), Width = 90 });
        diffGrid.Columns.Add(new System.Windows.Controls.DataGridTextColumn { Header = "컬럼", Binding = new System.Windows.Data.Binding(nameof(CellDiffItem.ColumnName)), Width = 220 });
        diffGrid.Columns.Add(new System.Windows.Controls.DataGridTextColumn { Header = "이전 값", Binding = new System.Windows.Data.Binding(nameof(CellDiffItem.BeforeValue)), Width = new System.Windows.Controls.DataGridLength(1, System.Windows.Controls.DataGridLengthUnitType.Star) });
        diffGrid.Columns.Add(new System.Windows.Controls.DataGridTextColumn { Header = "현재 값", Binding = new System.Windows.Data.Binding(nameof(CellDiffItem.AfterValue)), Width = new System.Windows.Controls.DataGridLength(1, System.Windows.Controls.DataGridLengthUnitType.Star) });

        System.Windows.Controls.Grid.SetRow(diffGrid, 1);
        grid.Children.Add(diffGrid);

        diffWindow.Content = grid;
        diffWindow.ShowDialog();
    }

    private void CheckinButton_Click(object sender, RoutedEventArgs e)
    {
        (bool statusOk, bool isCheckedOut, string statusMessage) = PerforceHelper.GetCheckoutStatus(_filePath);
        if (!statusOk)
        {
            MessageBox.Show(
                $"체크아웃 상태 확인에 실패했습니다.\n{statusMessage}",
                "체크인 불가",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            RefreshCheckoutState();
            return;
        }

        if (!isCheckedOut)
        {
            MessageBox.Show(
                "현재 파일은 체크아웃 상태가 아닙니다.\n체크아웃 후 체크인해 주세요.",
                "체크인 불가",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            RefreshCheckoutState();
            return;
        }

        string description = CheckinDescriptionTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(description))
        {
            EditorStatusTextBlock.Text = "체크인 설명을 입력해 주세요.";
            return;
        }

        SaveButton_Click(sender, e);

        string prefix = CheckinPrefixTextBox.Text.Trim();
        string fullDescription = string.IsNullOrWhiteSpace(prefix)
            ? description
            : $"{prefix} {description}";

        List<CellDiffItem> diffs = GetCurrentDiffItems();
        var confirmWindow = new CheckinConfirmWindow(fullDescription, diffs)
        {
            Owner = this
        };
        bool? confirmed = confirmWindow.ShowDialog();
        if (confirmed != true)
        {
            EditorStatusTextBlock.Text = "체크인이 취소되었습니다.";
            return;
        }

        string finalDescription = confirmWindow.FinalDescription;
        (bool success, string message) = PerforceHelper.Checkin(_filePath, finalDescription);
        EditorStatusTextBlock.Text = success
            ? "Perforce 체크인 완료"
            : $"체크인 실패: {message}";
        RefreshCheckoutState();
    }

    private void RevertButton_Click(object sender, RoutedEventArgs e)
    {
        List<CellDiffItem> diffs = GetCurrentDiffItems();
        var confirmWindow = new RevertConfirmWindow(diffs)
        {
            Owner = this
        };

        bool? confirmed = confirmWindow.ShowDialog();
        if (confirmed != true)
        {
            return;
        }

        (bool success, string message) = PerforceHelper.Revert(_filePath);
        if (!success)
        {
            EditorStatusTextBlock.Text = $"리버트 실패: {message}";
            RefreshCheckoutState();
            return;
        }

        EditorStatusTextBlock.Text = "리버트 완료";
        LoadSheetData();
        RefreshCheckoutState();
    }

    private void CheckinPrefixTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
    {
        SaveCheckinPrefix();
    }

    private void SheetDataGrid_LoadingRow(object sender, System.Windows.Controls.DataGridRowEventArgs e)
    {
        bool isTargetRow = e.Row.GetIndex() == _targetTableRowIndex;
        e.Row.Background = isTargetRow
            ? new SolidColorBrush(Color.FromRgb(255, 225, 110))
            : System.Windows.Media.Brushes.White;
        e.Row.FontWeight = isTargetRow ? FontWeights.SemiBold : FontWeights.Normal;

        if (isTargetRow && _highlightTableColumnIndexes.Count > 0)
        {
            e.Row.Dispatcher.BeginInvoke(new Action(() => HighlightCellsInRow(e.Row)), DispatcherPriority.Loaded);
        }
    }

    private void CaptureOriginalSnapshot()
    {
        _originalSnapshot.Clear();
        if (_table is null)
        {
            return;
        }

        for (int r = 0; r < _table.Rows.Count; r++)
        {
            for (int c = 0; c < _table.Columns.Count; c++)
            {
                _originalSnapshot[MakeSnapshotKey(r, c)] = _table.Rows[r][c]?.ToString() ?? string.Empty;
            }
        }
    }

    private List<CellDiffItem> GetCurrentDiffItems()
    {
        var result = new List<CellDiffItem>();
        if (_table is null)
        {
            return result;
        }

        for (int r = 0; r < _table.Rows.Count; r++)
        {
            for (int c = 0; c < _table.Columns.Count; c++)
            {
                string currentValue = _table.Rows[r][c]?.ToString() ?? string.Empty;
                string beforeValue = _originalSnapshot.TryGetValue(MakeSnapshotKey(r, c), out string? captured)
                    ? captured
                    : string.Empty;

                if (string.Equals(beforeValue, currentValue, StringComparison.Ordinal))
                {
                    continue;
                }

                int excelRow = r < _sheetRowNumbers.Count ? _sheetRowNumbers[r] : -1;
                string columnName = _table.Columns[c].ColumnName;

                result.Add(new CellDiffItem
                {
                    ExcelRow = excelRow,
                    ColumnName = columnName,
                    BeforeValue = beforeValue,
                    AfterValue = currentValue
                });
            }
        }

        return result;
    }

    private static string MakeSnapshotKey(int rowIndex, int columnIndex)
    {
        return $"{rowIndex}:{columnIndex}";
    }

    private void LoadSavedCheckinPrefix()
    {
        AppSettings settings = _settingsStore.Load();
        CheckinPrefixTextBox.Text = settings.CheckinPrefix ?? string.Empty;
    }

    private void SaveCheckinPrefix()
    {
        AppSettings settings = _settingsStore.Load();
        settings.CheckinPrefix = CheckinPrefixTextBox.Text.Trim();
        _settingsStore.Save(settings);
    }

    private void RefreshCheckoutState()
    {
        (bool success, bool isCheckedOut, string message) = PerforceHelper.GetCheckoutStatus(_filePath);
        CheckoutStateTextBlock.Text = success
            ? $"체크아웃 상태: {(isCheckedOut ? "ON" : "OFF")}"
            : $"체크아웃 상태 확인 실패: {message}";
    }

    private void HighlightCellsInRow(DataGridRow row)
    {
        DataGridCellsPresenter? presenter = FindVisualChild<DataGridCellsPresenter>(row);
        if (presenter is null)
        {
            return;
        }

        foreach (int colIndex in _highlightTableColumnIndexes)
        {
            DataGridCell? cell = GetCell(row, presenter, colIndex);
            if (cell is null)
            {
                continue;
            }

            cell.Background = new SolidColorBrush(Color.FromRgb(255, 199, 130));
            cell.BorderBrush = new SolidColorBrush(Color.FromRgb(230, 120, 20));
            cell.BorderThickness = new Thickness(1.5);
            cell.FontWeight = FontWeights.Bold;
        }
    }

    private static DataGridCell? GetCell(DataGridRow row, DataGridCellsPresenter presenter, int columnIndex)
    {
        DataGridCell? cell = presenter.ItemContainerGenerator.ContainerFromIndex(columnIndex) as DataGridCell;
        if (cell is not null)
        {
            return cell;
        }

        row.UpdateLayout();
        return presenter.ItemContainerGenerator.ContainerFromIndex(columnIndex) as DataGridCell;
    }

    private static T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
    {
        int childCount = VisualTreeHelper.GetChildrenCount(parent);
        for (int i = 0; i < childCount; i++)
        {
            DependencyObject child = VisualTreeHelper.GetChild(parent, i);
            if (child is T target)
            {
                return target;
            }

            T? nested = FindVisualChild<T>(child);
            if (nested is not null)
            {
                return nested;
            }
        }

        return null;
    }

    private static int GetColumnIndexFromCellAddress(string? cellAddress)
    {
        if (string.IsNullOrWhiteSpace(cellAddress))
        {
            return -1;
        }

        string letters = new string(cellAddress.TakeWhile(char.IsLetter).ToArray()).ToUpperInvariant();
        if (string.IsNullOrEmpty(letters))
        {
            return -1;
        }

        int column = 0;
        foreach (char c in letters)
        {
            column = (column * 26) + (c - 'A' + 1);
        }

        return column;
    }
}

public sealed class CellDiffItem
{
    public int ExcelRow { get; set; }
    public string ColumnName { get; set; } = string.Empty;
    public string BeforeValue { get; set; } = string.Empty;
    public string AfterValue { get; set; } = string.Empty;
}
