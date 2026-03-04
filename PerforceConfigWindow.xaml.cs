using System.Windows;

namespace ExcelFinder;

public partial class PerforceConfigWindow : Window
{
    public PerforceConfigWindow(PerforceClientSettings? initialSettings = null)
    {
        InitializeComponent();
        Loaded += PerforceConfigWindow_Loaded;

        if (initialSettings is null)
        {
            return;
        }

        ClientTextBox.Text = initialSettings.Client;
        HostTextBox.Text = initialSettings.Host;
        RootTextBox.Text = initialSettings.Root;
        StreamTextBox.Text = initialSettings.Stream;
    }

    private async void PerforceConfigWindow_Loaded(object sender, RoutedEventArgs e)
    {
        await RefreshP4InfoAsync();
    }

    public PerforceClientSettings CurrentSettings => new()
    {
        Client = ClientTextBox.Text.Trim(),
        Host = HostTextBox.Text.Trim(),
        Root = RootTextBox.Text.Trim(),
        Stream = StreamTextBox.Text.Trim()
    };

    private async void ApplyButton_Click(object sender, RoutedEventArgs e)
    {
        PerforceClientSettings settings = CurrentSettings;
        if (string.IsNullOrWhiteSpace(settings.Client))
        {
            StatusTextBlock.Text = "Client 값을 입력해 주세요.";
            return;
        }

        StatusTextBlock.Text = "Perforce 정보 적용 중...";

        (bool success, string message) = await Task.Run(() => PerforceHelper.ApplyClientSettings(settings));
        StatusTextBlock.Text = success ? "적용 완료" : $"적용 실패: {message}";
        await RefreshP4InfoAsync();

        if (success)
        {
            DialogResult = true;
            Close();
        }
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private async Task RefreshP4InfoAsync()
    {
        (bool success, string info) = await Task.Run(PerforceHelper.GetInfoRaw);
        P4InfoTextBox.Text = info;
        if (!success)
        {
            StatusTextBlock.Text = $"p4 info 조회 실패: {info}";
        }
    }
}
