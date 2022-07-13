using IronXL;

namespace XLReporting.Reporting;

public interface IReportGenerator
{
    public string GenerateReport();

    public WorkBook CreateTemplate();

    public WorkBook LoadTemplate();
}