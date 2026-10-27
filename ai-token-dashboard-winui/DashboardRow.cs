namespace AiTokenDashboard;

public sealed class DashboardRow
{
    public string Name { get; set; } = string.Empty;
    public string Quota { get; set; } = string.Empty;
    public double QuotaPercent { get; set; }
    public string MonthlyTokens { get; set; } = "─";
    public string Cost { get; set; } = "─";
    public string Plan { get; set; } = "─";
    public string Reset { get; set; } = "─";
}
