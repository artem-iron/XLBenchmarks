namespace XLReporting.Configuration;

public class AppConfig : IAppConfig
{
    public string? Environment { get; set; }
    public string? LicenseKey { get; set; }
    public string ReportsFolder { get; set; } = "Reports";
    public int ChartWidth { get; set; } = 11;
    public int ChartHeight { get; set; } = 24;
    public int ContendersNumber { get; set; } = 3;
    public int TimeTableStartingRow { get; set; } = 27;

    public string? ChartTitle { get; set; }
    public string[]? TestList { get; set; }

    public int DateCellsNumber { get; set; } = 80000;
    public int RandomCellsRowNumber { get; set; } = 20000;
    public int StyleChangeRowNumber { get; set; } = 300;

    public string ResultsFolderName { get; set; } = "Results";
    public string RandomCellsFileNameTemplate { get; set; } = "{0}\\{1}_RandomCells.xlsx";
    public string DateCellsFileNameTemplate { get; set; } = "{0}\\{1}_DateCells.xlsx";
    public string StyleChangeFileNameTemplate { get; set; } = "{0}\\{1}_StyleChange.xlsx";
    public string LoadingLargeFileFileNameTemplate { get; set; } = "{0}\\{1}_LoadingBigFile.xlsx";
    public string CellValue { get; set; } = "Cell";
    public string LargeFileName { get; set; } = "LoadingTestFiles\\LoadingTest.xlsx";
}