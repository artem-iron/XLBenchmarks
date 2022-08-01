using IronXL;

namespace XLBenchmarks.Reporting;

public interface IReportGenerator
{
    public string GenerateReport();

    public WorkBook CreateTemplate();

    public WorkBook LoadTemplate();

    public void FillReport(WorkBook report);
}