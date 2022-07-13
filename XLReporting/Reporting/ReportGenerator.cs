using System.Reflection;
using IronXL;
using IronXL.Drawing.Charts;
using IronXL.Formatting;
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

        var fileName = Path.Combine(path, $"Report_{DateTime.Now:yyyy-MM-d_HH-mm-ss}.xlsx");

        var report = WorkBook.Create(ExcelFileFormat.XLSX);
        _ = report.DefaultWorkSheet;
        report.SaveAs(fileName);

        return fileName;
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
        var secondsInADay = 60 * 60 * 24;

        foreach (var cell in sheet[seriesRowAddress])
        {
            cell.Value = (double)rnd.Next(25, 100) / secondsInADay;
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
            cell.Value = $"Test_{cell.ColumnIndex}";
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