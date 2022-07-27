using IronXL;
using IronXL.Drawing.Charts;
using IronXL.Formatting;
using System.Reflection;
using XLReporting.Configuration;
using XLReporting.BenchmarkRunners;

namespace XLReporting.Reporting;

public class ReportGenerator : IReportGenerator
{
    private readonly IAppConfig _appConfig;
    private readonly string headerRowAddress;

    public ReportGenerator(IAppConfig appConfig)
    {
        _appConfig = appConfig;
        headerRowAddress = $"B{_appConfig.TimeTableStartingRow}:K{_appConfig.TimeTableStartingRow}";
    }

    public string GenerateReport()
    {
        CreateReportsFolder();

        var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        if (path == null)
        {
            return "";
        }

        var reportName = Path.Combine(path, $"{_appConfig.ReportsFolder}\\Report_{DateTime.Now:yyyy-MM-d_HH-mm-ss}.xlsx");

        var report = LoadTemplate();

        FillReport(report);

        report.SaveAs(reportName);

        return reportName;
    }

    public WorkBook CreateTemplate()
    {
        var template = WorkBook.Create(ExcelFileFormat.XLSX);
        var sheet = template.DefaultWorkSheet;

        PutInMockData(sheet);

        AddChart(sheet);

        FormatTimeTable(sheet);

        template.SaveAs("template.xlsx");
        template = WorkBook.Load("template.xlsx");

        return template;
    }

    public WorkBook LoadTemplate()
    {
        if (File.Exists(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\template.xlsx"))
        {
            return WorkBook.Load("template.xlsx");
        }

        return CreateTemplate();
    }

    public void FillReport(WorkBook report)
    {
        Dictionary<string, TimeSpan[]> timeTableData = new()
        {
            { "Current IronXL", GetCurrentIronXLBenchmarkData() },
            { "Previous IronXL", GetPreviousIronXLBenchmarkData() },
            { "Aspose", GetAsposeBenchmarkData() },
            { "NPOI", GetNpoiBenchmarkData() },
        };

        _appConfig.ContendersNumber = timeTableData.Count;

        var sheet = report.DefaultWorkSheet;

        FillHeader(sheet, headerRowAddress);

        var i = 0;

        foreach (var contender in timeTableData.Keys)
        {
            i++;

            var times = timeTableData[contender];

            FillRow(sheet, i, contender, times);
        }

        UpdateChart(sheet);
    }

    private void CreateReportsFolder()
    {
        var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        if (path == null)
        {
            return;
        }

        var reportsFolder = Path.Combine(path, _appConfig.ReportsFolder);

        if (!Directory.Exists(reportsFolder))
        {
            Directory.CreateDirectory(reportsFolder);
        }
    }

    private void FillHeader(WorkSheet sheet, string headerRowAddress)
    {
        string[] benchmarkList = _appConfig.BenchmarkList ?? new string[] { "couldn't get benchmark list from config" };

        var i = 0;

        foreach (var cell in sheet[headerRowAddress])
        {
            cell.Value = benchmarkList[i];

            i++;
        }
    }

    private TimeSpan[] GetAsposeBenchmarkData()
    {
        return new AsposeBenchmarkRunner(_appConfig).RunBenchmarks();
    }

    private TimeSpan[] GetPreviousIronXLBenchmarkData()
    {
        return new PreviousIronXLBenchmarkRunner(_appConfig).RunBenchmarks();
    }

    private TimeSpan[] GetCurrentIronXLBenchmarkData()
    {
        return new CurrentIronXLBenchmarkRunner(_appConfig).RunBenchmarks();
    }

    private TimeSpan[] GetNpoiBenchmarkData()
    {
        return new NpoiBenchmarkRunner(_appConfig).RunBenchmarks();
    }

    private void FillRow(WorkSheet sheet, int i, string contender, TimeSpan[] times)
    {
        var seriesRowNumber = _appConfig.TimeTableStartingRow + i;
        var seriesRowAddress = $"B{seriesRowNumber}:K{seriesRowNumber}";

        PutInSeriesData(sheet, seriesRowAddress, times);

        sheet[$"A{seriesRowNumber}"].Value = contender;
    }

    private void FormatTimeTable(WorkSheet sheet)
    {
        for (var i = 1; i <= _appConfig.ContendersNumber; i++)
        {
            var seriesRowNumber = _appConfig.TimeTableStartingRow + i;
            var seriesRowAddress = $"B{seriesRowNumber}:K{seriesRowNumber}";

            FormatRow(sheet, seriesRowAddress);
        }
    }

    private void PutInMockData(WorkSheet sheet)
    {
        PutInMockHeaderData(sheet, headerRowAddress);

        PutInMockTimeTableData(sheet);
    }

    private void PutInMockTimeTableData(WorkSheet sheet)
    {
        for (var i = 1; i <= _appConfig.ContendersNumber; i++)
        {
            var seriesRowNumber = _appConfig.TimeTableStartingRow + i;
            var seriesRowAddress = $"B{seriesRowNumber}:K{seriesRowNumber}";

            PutInMockSeriesData(sheet, seriesRowAddress);

            sheet[$"A{seriesRowNumber}"].Value = $"Contender_{seriesRowNumber}";
        }
    }

    private static void PutInMockSeriesData(WorkSheet sheet, string seriesRowAddress)
    {
        var rnd = new Random();
        var times = new TimeSpan[10];

        for (int i = 0; i < times.Length; i++)
        {
            times[i] = TimeSpan.FromSeconds(rnd.Next(25, 100));
        }

        PutInSeriesData(sheet, seriesRowAddress, times);
    }

    private static void PutInSeriesData(WorkSheet sheet, string seriesRowAddress, TimeSpan[] times)
    {
        var secondsInADay = 60 * 60 * 24;

        var i = 0;

        foreach (var cell in sheet[seriesRowAddress])
        {
            cell.Value = times[i].TotalSeconds / secondsInADay;

            i++;
        }
    }

    private static void FormatRow(WorkSheet sheet, string rowAddress)
    {
        sheet[rowAddress].FormatString = BuiltinFormats.Duration3;
    }

    private static void PutInMockHeaderData(WorkSheet sheet, string headerRowAddress)
    {
        foreach (var cell in sheet[headerRowAddress])
        {
            cell.Value = $"Mock_Benchmark_{cell.ColumnIndex}";
        }
    }

    private void AddChart(WorkSheet sheet)
    {
        var chart = sheet.CreateChart(ChartType.Bar, 0, 0, _appConfig.ChartHeight, _appConfig.ChartWidth);

        for (var i = 1; i <= _appConfig.ContendersNumber; i++)
        {
            var seriesRowNumber = _appConfig.TimeTableStartingRow + i;
            var seriesRowAddress = $"B{seriesRowNumber}:K{seriesRowNumber}";

            var range = sheet[seriesRowAddress];
            range.FormatString = BuiltinFormats.Number0;

            var series = chart.AddSeries(seriesRowAddress, headerRowAddress);
            series.Title = sheet[$"A{seriesRowNumber}"].StringValue;
        }

        chart.SetTitle(_appConfig.ChartTitle);
        chart.SetLegendPosition(LegendPosition.Bottom);
        chart.Plot();
    }

    private static void RemoveChart(WorkSheet sheet)
    {
        var chart = sheet.Charts.FirstOrDefault();

        if (chart != null)
        {
            sheet.Charts.Remove(chart);
        }
    }

    private void UpdateChart(WorkSheet sheet)
    {
        RemoveChart(sheet);

        AddChart(sheet);

        FormatTimeTable(sheet);
    }
}