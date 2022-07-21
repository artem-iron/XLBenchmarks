using System.Reflection;
using Microsoft.Extensions.Configuration;

namespace XLReporting.Tests;

public abstract class TestsBase
{
    protected TestsBase()
    {
        var builder = new ConfigurationBuilder()
            .AddJsonFile($"appsettings.json", true, true)
            .AddUserSecrets(Assembly.GetExecutingAssembly(), true)
            .AddEnvironmentVariables();

        var configurationRoot = builder.Build();

        IronXL.License.LicenseKey = configurationRoot.GetSection("LicenseKey").Value;
        IronXLOld.License.LicenseKey = configurationRoot.GetSection("LicenseKey").Value;

        Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.InvariantCulture;
    }
}