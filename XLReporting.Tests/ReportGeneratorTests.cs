using Moq;
using XLReporting.Configuration;
using XLReporting.Reporting;

namespace XLReporting.Tests;

[TestClass]
public class ReportGeneratorTests : TestsBase
{
    private readonly ReportGenerator _reportGenerator;
    private readonly Mock<IAppConfig> _appConfig;

    public ReportGeneratorTests() : base()
    {
        _appConfig = new Mock<IAppConfig>();
        _appConfig.SetupAllProperties();
        _appConfig.Object.ChartTitle = "Chart Title";
        _appConfig.Object.ChartHeight = 24;
        _appConfig.Object.ChartWidth = 11;
        _appConfig.Object.ContendersNumber = 3;
        _appConfig.Object.TimeTableStartingRow = 27;
        _appConfig.Object.TestList = new string[] {
            "Test1", "Test2", "Test3", "Test4",
            "Test5", "Test6", "Test7", "Test8",
            "Test9", "Test10"
        };

        _reportGenerator = new ReportGenerator(_appConfig.Object);
    }

    [TestMethod]
    public void CreateReport_ReportIsCreated()
    {
        var reportName = _reportGenerator.GenerateReport();
        var fileExists = File.Exists(reportName);

        Assert.IsTrue(fileExists);
    }

    [TestMethod]
    public void CreateTemplate_TemplateIsCreated()
    {
        var template = _reportGenerator.CreateTemplate();

        Assert.IsNotNull(template);
    }

    [TestMethod]
    public void LoadTemplate_TemplateIsLoadedOrCreated()
    {
        var template = _reportGenerator.LoadTemplate();

        Assert.IsNotNull(template);
    }

    [TestMethod]
    public void FillReport_ReportIsFilled()
    {
        var report = _reportGenerator.LoadTemplate();

        _reportGenerator.FillReport(report);

        Assert.IsFalse(report.DefaultWorkSheet.FilledCells.Any(c => c.StringValue.Contains("Contender_")));
        Assert.IsFalse(report.DefaultWorkSheet.FilledCells.Any(c => c.StringValue.Contains("Mock_Test_")));
    }
}