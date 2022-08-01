using XLBenchmarks.Reporting;

namespace XLBenchmarks.Tests;

[TestClass]
public class ReportGeneratorTests : TestsBase
{
    private readonly ReportGenerator _reportGenerator;

    public ReportGeneratorTests() : base()
    {
        _reportGenerator = new ReportGenerator(_appConfig);
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
        Assert.IsFalse(report.DefaultWorkSheet.FilledCells.Any(c => c.StringValue.Contains("Mock_Benchmark_")));
    }
}