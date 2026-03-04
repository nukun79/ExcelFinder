using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;

namespace ExcelFinder;

public partial class HistoryWindow : Window
{
    private readonly string _filePath;
    private string _depotPath = string.Empty;
    private readonly ObservableCollection<PerforceFileHistoryEntry> _entries = [];

    public HistoryWindow(string filePath)
    {
        InitializeComponent();

        _filePath = filePath;
        HeaderTextBlock.Text = $"History of File: {_filePath}";
        HistoryDataGrid.ItemsSource = _entries;

        Loaded += HistoryWindow_Loaded;
    }

    private async void HistoryWindow_Loaded(object sender, RoutedEventArgs e)
    {
        await LoadHistoryAsync();
    }

    private async Task LoadHistoryAsync()
    {
        if (!File.Exists(_filePath))
        {
            StatusTextBlock.Text = "파일을 찾을 수 없습니다.";
            return;
        }

        StatusTextBlock.Text = "히스토리 조회 중...";

        (bool success, string message, List<PerforceFileHistoryEntry> entries) = await Task.Run(() => PerforceHelper.GetFileHistory(_filePath));
        _entries.Clear();

        if (!success)
        {
            StatusTextBlock.Text = $"히스토리 조회 실패: {message}";
            return;
        }

        foreach (PerforceFileHistoryEntry entry in entries)
        {
            _entries.Add(entry);
        }

        (bool depotOk, string depotMsg, string depotPath) = await Task.Run(() => PerforceHelper.GetDepotPathForLocalFile(_filePath));
        _depotPath = depotOk ? depotPath : string.Empty;

        StatusTextBlock.Text = entries.Count > 0
            ? $"히스토리 {entries.Count}건{(depotOk ? string.Empty : $" / depot 경로 조회 실패: {depotMsg}")}"
            : "히스토리가 없습니다.";
    }

    private void HistoryDataGrid_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        DependencyObject? source = e.OriginalSource as DependencyObject;
        if (source is null)
        {
            return;
        }

        DataGridRow? row = FindVisualParent<DataGridRow>(source);
        if (row?.Item is PerforceFileHistoryEntry)
        {
            row.IsSelected = true;
            HistoryDataGrid.SelectedItem = row.Item;
        }
    }

    private void OpenDiffMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (HistoryDataGrid.SelectedItem is not PerforceFileHistoryEntry selected)
        {
            StatusTextBlock.Text = "선택된 히스토리 항목이 없습니다.";
            return;
        }

        if (string.IsNullOrWhiteSpace(_depotPath))
        {
            StatusTextBlock.Text = "depot 경로 정보가 없어 Diff를 열 수 없습니다.";
            return;
        }

        if (!int.TryParse(selected.Revision, out int revision) || revision <= 1)
        {
            StatusTextBlock.Text = "Diff는 2번 이상 리비전에서만 가능합니다.";
            return;
        }

        var diffWindow = new HistoryDiffWindow(_depotPath, revision)
        {
            Owner = this
        };
        diffWindow.ShowDialog();
    }

    private async void DownloadMenuItem_Click(object sender, RoutedEventArgs e)
    {
        if (HistoryDataGrid.SelectedItem is not PerforceFileHistoryEntry selected)
        {
            StatusTextBlock.Text = "선택된 히스토리 항목이 없습니다.";
            return;
        }

        if (string.IsNullOrWhiteSpace(_depotPath))
        {
            StatusTextBlock.Text = "depot 경로 정보가 없어 다운로드할 수 없습니다.";
            return;
        }

        string extension = Path.GetExtension(_filePath);
        string baseName = Path.GetFileNameWithoutExtension(_filePath);
        string initialDir = Path.GetDirectoryName(_filePath) ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        string proposedName = $"{baseName}_rev{selected.Revision}{extension}";

        var dialog = new SaveFileDialog
        {
            Title = "리비전 파일 다운로드",
            InitialDirectory = Directory.Exists(initialDir) ? initialDir : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            FileName = proposedName,
            Filter = "Excel Files|*.xlsx;*.xlsm|All Files|*.*"
        };

        bool? result = dialog.ShowDialog(this);
        if (result != true)
        {
            StatusTextBlock.Text = "다운로드가 취소되었습니다.";
            return;
        }

        string spec = $"{_depotPath}#{selected.Revision}";
        StatusTextBlock.Text = $"다운로드 중... ({spec})";

        (bool success, string message) = await Task.Run(() => PerforceHelper.ExportDepotRevisionToFile(spec, dialog.FileName));
        StatusTextBlock.Text = success
            ? $"다운로드 완료: {dialog.FileName}"
            : $"다운로드 실패: {message}";
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
