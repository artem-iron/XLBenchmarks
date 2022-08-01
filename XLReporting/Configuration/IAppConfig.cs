namespace XLBenchmarks.Configuration;

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
    public string[]? BenchmarkList { get; set; }

    public int DateCellsNumber { get; set; }
    public int RandomCellsRowNumber { get; set; }
    public int StyleChangeRowNumber { get; set; }
    public int GenerateFormulasRowNumber { get; set; }

    public string ResultsFolderName { get; set; }
    public string RandomCellsFileNameTemplate { get; set; }
    public string DateCellsFileNameTemplate { get; set; }
    public string StyleChangeFileNameTemplate { get; set; }
    public string LoadingLargeFileFileNameTemplate { get; set; }
    public string GenerateFormulasFileNameTemplate { get; set; }
    public string CellValue { get; set; }
    public string LargeFileName { get; set; }
}