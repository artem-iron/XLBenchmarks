namespace XLReporting.Configuration;

public interface IAppConfig
{
    public string? Environment { get; set; }
    public string? LicenseKey { get; }
}