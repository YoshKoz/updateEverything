using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.UI.Dispatching;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Windows.Storage;
using Windows.UI;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace AiTokenDashboard;

public sealed partial class MainPage : Page
{
    private static readonly Regex TailPattern = new(
        @"(?<plan>[\w.+-]{0,14})\s+(?<reset>soon|in\s+\d+\w+|─)$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex CostPattern = new(
        @"(?<cost>€[\d,.]+/mo|sub|─)\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex PercentPattern = new(@"(?<percent>\d{1,3}(?:\.\d+)?)%", RegexOptions.Compiled);
    private readonly DispatcherTimer _refreshTimer = new() { Interval = TimeSpan.FromMinutes(1) };

    public ObservableCollection<DashboardRow> Rows { get; } = new();

    public MainPage()
    {
        InitializeComponent();
        ExecutableBox.Text = GetDefaultExecutablePath();
        _refreshTimer.Tick += (_, _) => _ = RefreshSnapshotAsync();
    }

    private async void Page_Loaded(object sender, RoutedEventArgs e)
    {
        _refreshTimer.Start();
        await RefreshSnapshotAsync();
    }

    private async void Refresh_Click(object sender, RoutedEventArgs e) => await RefreshSnapshotAsync();

    private async Task RefreshSnapshotAsync()
    {
        if (LoadingRing.IsActive)
        {
            return;
        }

        var executable = ExecutableBox.Text.Trim().Trim('"');
        if (string.IsNullOrWhiteSpace(executable))
        {
            ShowConnection(false, "Select the local dashboard executable");
            return;
        }

        LoadingRing.IsActive = true;
        RefreshButton.IsEnabled = false;
        ErrorText.Text = "Reading local AI tool usage...";
        try
        {
            if (!File.Exists(executable))
            {
                throw new FileNotFoundException("The local dashboard executable was not found.", executable);
            }

            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
            var startInfo = new ProcessStartInfo(executable)
            {
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WorkingDirectory = Path.GetDirectoryName(executable) ?? Environment.CurrentDirectory
            };
            startInfo.ArgumentList.Add("--once");
            startInfo.ArgumentList.Add("--no-remote");
            startInfo.ArgumentList.Add("--log-level");
            startInfo.ArgumentList.Add("error");
            var cacheDirectory = Path.Combine(ApplicationData.Current.LocalCacheFolder.Path, "ai-token-dashboard");
            Directory.CreateDirectory(cacheDirectory);
            startInfo.Environment["AI_TOKEN_DASHBOARD_CACHE"] = cacheDirectory;

            using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Could not start the dashboard backend.");
            var stdoutTask = process.StandardOutput.ReadToEndAsync(timeout.Token);
            var stderrTask = process.StandardError.ReadToEndAsync(timeout.Token);
            try
            {
                await process.WaitForExitAsync(timeout.Token);
            }
            catch (OperationCanceledException)
            {
                process.Kill(entireProcessTree: true);
                throw new TimeoutException("The dashboard backend did not return within 30 seconds.");
            }

            var stdout = await stdoutTask;
            var stderr = await stderrTask;
            if (process.ExitCode != 0)
            {
                throw new InvalidOperationException(string.IsNullOrWhiteSpace(stderr)
                    ? $"The dashboard backend exited with code {process.ExitCode}."
                    : stderr.Trim());
            }

            var rows = ParseSnapshot(stdout);
            Rows.Clear();
            foreach (var row in rows)
            {
                Rows.Add(row);
            }

            UpdateSummary(rows);
            EmptyPanel.Visibility = rows.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            ErrorText.Text = rows.Count == 0 ? "The local dashboard returned no rows." : string.Empty;
            UpdatedText.Text = "Updated " + DateTime.Now.ToString("HH:mm:ss") + "  |  Next refresh in 1 minute";
            ShowConnection(true, "Local data ready");
        }
        catch (Exception ex)
        {
            Rows.Clear();
            ToolsValue.Text = "-";
            TokensValue.Text = "-";
            CostValue.Text = "-";
            EmptyPanel.Visibility = Visibility.Visible;
            ErrorText.Text = ex.Message;
            UpdatedText.Text = "Could not refresh. Check the local executable and its configuration.";
            ShowConnection(false, "Local backend unavailable");
        }
        finally
        {
            LoadingRing.IsActive = false;
            RefreshButton.IsEnabled = true;
        }
    }

    private static string GetDefaultExecutablePath()
    {
        return Path.Combine(AppContext.BaseDirectory, "ai-token-dashboard.exe");
    }

    private static List<DashboardRow> ParseSnapshot(string output)
    {
        var rows = new List<DashboardRow>();
        foreach (var rawLine in output.Split('\n'))
        {
            var line = rawLine.TrimEnd('\r');
            if (string.IsNullOrWhiteSpace(line) || line.StartsWith("Tool ", StringComparison.Ordinal) ||
                line.StartsWith("──", StringComparison.Ordinal) || line.All(c => c == '─'))
            {
                continue;
            }

            var tail = TailPattern.Match(line);
            if (tail.Success)
            {
                var beforePlan = line[..tail.Index].TrimEnd();
                var cost = CostPattern.Match(beforePlan);
                var tokenStart = cost.Success ? cost.Index - 13 : -1;
                if (cost.Success && tokenStart >= 17)
                {
                    var quota = line.Substring(17, tokenStart - 17).Trim();
                    var tokenWidth = Math.Min(12, line.Length - tokenStart);
                    var tokenText = line.Substring(tokenStart, tokenWidth).Trim();
                    var percentMatch = PercentPattern.Match(quota);
                    var percent = percentMatch.Success && double.TryParse(
                        percentMatch.Groups["percent"].Value,
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture,
                        out var parsedPercent)
                        ? Math.Clamp(parsedPercent, 0, 100)
                        : 0;
                    rows.Add(new DashboardRow
                    {
                        Name = line[..Math.Min(16, line.Length)].Trim(),
                        Quota = quota,
                        QuotaPercent = percent,
                        MonthlyTokens = tokenText,
                        Cost = cost.Groups["cost"].Value,
                        Plan = tail.Groups["plan"].Value.Trim(),
                        Reset = tail.Groups["reset"].Value.Trim()
                    });
                    continue;
                }
            }

            var name = line[..Math.Min(16, line.Length)].Trim();
            var status = line[Math.Min(16, line.Length)..].Trim();
            if (name.Length > 0 && status.Length > 0)
            {
                rows.Add(new DashboardRow
                {
                    Name = name,
                    Quota = status,
                    MonthlyTokens = "─",
                    Cost = "─",
                    Plan = "─",
                    Reset = "─"
                });
            }
        }

        return rows;
    }

    private void UpdateSummary(IReadOnlyCollection<DashboardRow> rows)
    {
        ToolsValue.Text = rows.Count.ToString(CultureInfo.InvariantCulture);
        var monthlyTokens = rows.Sum(row => ParseCompactNumber(row.MonthlyTokens));
        TokensValue.Text = monthlyTokens == 0 ? "-" : FormatCompactNumber(monthlyTokens);
        var monthlyCost = rows.Sum(row => ParseMonthlyCost(row.Cost));
        CostValue.Text = monthlyCost == 0 ? "-" : $"€{monthlyCost:N2}";
    }

    private static decimal ParseCompactNumber(string value)
    {
        var trimmed = value.Trim();
        if (trimmed.Length == 0 || trimmed == "─")
        {
            return 0;
        }

        var multiplier = trimmed[^1] switch
        {
            'K' => 1_000m,
            'M' => 1_000_000m,
            'B' => 1_000_000_000m,
            _ => 1m
        };
        var number = multiplier == 1m ? trimmed : trimmed[..^1];
        return decimal.TryParse(number, NumberStyles.Number, CultureInfo.InvariantCulture, out var parsed)
            ? parsed * multiplier
            : 0;
    }

    private static decimal ParseMonthlyCost(string value)
    {
        if (!value.StartsWith('€') || !value.EndsWith("/mo", StringComparison.Ordinal))
        {
            return 0;
        }

        return decimal.TryParse(value[1..^3], NumberStyles.Number, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : 0;
    }

    private static string FormatCompactNumber(decimal value) => value switch
    {
        >= 1_000_000_000m => $"{value / 1_000_000_000m:0.0}B",
        >= 1_000_000m => $"{value / 1_000_000m:0.0}M",
        >= 1_000m => $"{value / 1_000m:0.0}K",
        _ => $"{value:N0}"
    };

    private void ShowConnection(bool connected, string text)
    {
        ConnectionText.Text = text;
        ConnectionDot.Fill = new SolidColorBrush(connected ? Color.FromArgb(255, 98, 214, 199) : Color.FromArgb(255, 230, 184, 92));
    }
}
