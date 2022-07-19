using IronXL;
using IronXL.Drawing.Charts;
using IronXL.Formatting;
using System.Reflection;
using XLReporting.Configuration;

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
        var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

        if (path == null)
        {
            return "";
        }

        var reportName = Path.Combine(path, $"Report_{DateTime.Now:yyyy-MM-d_HH-mm-ss}.xlsx");

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
            { "Current IronXL", GetCurrentIronXLTestData() },
            { "Previous IronXL", GetPreviousIronXLTestData() },
            { "Aspose", GetAsposeTestData() },
        };

        var sheet = report.DefaultWorkSheet;

        FillHeader(sheet, headerRowAddress);
        
        var i = 0;
        
        foreach (var contender in timeTableData.Keys)
        {
            i++;
            
            var times = timeTableData[contender];

            FillRow(sheet, i, contender, times);
        }
    }

    private void FillHeader(WorkSheet sheet, string headerRowAddress)
    {
        string[] testList = _appConfig.TestList;

        var i = 0;

        foreach (var cell in sheet[headerRowAddress])
        {
            cell.Value = testList[i];

            i++;
        }
    }

    private static TimeSpan[] GetAsposeTestData()
    {
        var rnd = new Random();
        var times = new TimeSpan[10];

        for (int i = 0; i < times.Length; i++)
        {
            times[i] = TimeSpan.FromSeconds(rnd.Next(25, 100));
        }

        return times;
    }

    private static TimeSpan[] GetPreviousIronXLTestData()
    {
        return PreviousIronXL.TestRunner.RunTests();
    }

    private static TimeSpan[] GetCurrentIronXLTestData()
    {
        return CurrentIronXL.TestRunner.RunTests();
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
            cell.Value = $"Mock_Test_{cell.ColumnIndex}";
        }
    }

    private void AddChart(WorkSheet sheet)
    {
        var chart = sheet.CreateChart(ChartType.Bar, 0, 0, _appConfig.ChartHeight, _appConfig.ChartWidth);

        for (var i = 1; i <= _appConfig.ContendersNumber; i++)
        {
            var seriesRowNumber = _appConfig.TimeTableStartingRow + i;
            var seriesRowAddress = $"B{seriesRowNumber}:K{seriesRowNumber}";

            var series = chart.AddSeries(seriesRowAddress, headerRowAddress);
            series.Title = sheet[$"A{seriesRowNumber}"].StringValue;
        }

        chart.SetTitle(_appConfig.ChartTitle);
        chart.SetLegendPosition(LegendPosition.Bottom);
        chart.Plot();
    }
}