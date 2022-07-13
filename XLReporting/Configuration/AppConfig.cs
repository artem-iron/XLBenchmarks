namespace XLReporting.Configuration;

public class AppConfig : IAppConfig
{
    public string? Environment { get; set; }
    public string? LicenseKey { get; set; }
    public int ChartWidth { get; set; } = 11;
    public int ChartHeight { get; set; } = 24;
    public int ContendersNumber { get; set; } = 3;
    public int TimeTableStartingRow { get; set; } = 27;
    public string? ChartTitle { get; set; }
}