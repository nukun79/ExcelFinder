using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace ExcelFinder;

public partial class CheckinConfirmWindow : Window
{
    public string FinalDescription => DescriptionTextBox.Text.Trim();

    public CheckinConfirmWindow(string description, IReadOnlyList<CellDiffItem> diffs)
    {
        InitializeComponent();

        DescriptionTextBox.Text = description;
        DiffDataGrid.ItemsSource = diffs;
        SummaryTextBlock.Text = diffs.Count > 0
            ? $"변경 건수: {diffs.Count}"
            : "변경점이 없습니다. 그래도 체크인을 진행할 수 있습니다.";
    }

    private void ConfirmButton_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(FinalDescription))
        {
            SummaryTextBlock.Text = "Description을 입력해 주세요.";
            return;
        }

        DialogResult = true;
        Close();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
