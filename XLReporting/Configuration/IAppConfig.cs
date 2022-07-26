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
    int DateCellsNumber { get; set; }
    int RandomCellsRowNumber { get; set; }
    int StyleChangeRowNumber { get; set; }
    string ResultsFolderName { get; set; }
    string RandomCellsFileNameTemplate { get; set; }
    string DateCellsFileNameTemplate { get; set; }
    string StyleChangeFileNameTemplate { get; set; }
    string LoadingLargeFileFileNameTemplate { get; set; }
    string CellValue { get; set; }
    string LargeFileName { get; set; }
}