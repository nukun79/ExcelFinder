using System.Collections.Generic;
using System.Windows;

namespace ExcelFinder;

public partial class RevertConfirmWindow : Window
{
    public RevertConfirmWindow(IReadOnlyList<CellDiffItem> diffs)
    {
        InitializeComponent();

        DiffDataGrid.ItemsSource = diffs;
        SummaryTextBlock.Text = diffs.Count > 0
            ? $"변경 건수: {diffs.Count}"
            : "변경점이 없습니다. 그래도 리버트를 진행할 수 있습니다.";
    }

    private void ConfirmButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}