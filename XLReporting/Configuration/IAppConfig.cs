namespace XLReporting.Configuration;

public interface IAppConfig
{
    public string? Environment { get; set; }
    public string? LicenseKey { get; set; }
    public string ReportsFolder { get; set; }
    public int ChartWidth { get; set; }
    public int ChartHeight { get; set; }
    public int ContendersNumber { get; set; }
    public int TimeTableStartingRow { get; set; }
    public string? ChartTitle { get; set; }

    public string[] TestList { get; set; }
}