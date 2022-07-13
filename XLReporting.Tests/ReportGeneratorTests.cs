using Moq;
using XLReporting.Configuration;
using XLReporting.Reporting;

namespace XLReporting.Tests;

[TestClass]
public class ReportGeneratorTests : TestsBase
{
    private readonly ReportGenerator _reportGenerator;
    private Mock<IAppConfig> _appConfig;

    public ReportGeneratorTests() : base()
    {
        _appConfig = new Mock<IAppConfig>();
        _appConfig.SetupAllProperties();
        _appConfig.Object.ChartTitle = "Chart Title";
        _appConfig.Object.ChartHeight = 24;
        _appConfig.Object.ChartWidth = 11;
        _appConfig.Object.ContendersNumber = 3;
        _appConfig.Object.TimeTableStartingRow = 27;

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
        string template = null;//_reportGenerator.LoadTemplate();
        
        Assert.IsNotNull(template);
    }
    
    
}