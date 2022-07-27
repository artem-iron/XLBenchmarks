using System.Reflection;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using XLReporting.Configuration;

namespace XLReporting.Tests;

public abstract class TestsBase
{
    protected readonly IAppConfig _appConfig;

    protected TestsBase()
    {
        var builder = new ConfigurationBuilder()
            .AddJsonFile($"appsettings.json", true, true)
            .AddUserSecrets(Assembly.GetExecutingAssembly(), true)
            .AddEnvironmentVariables();

        var configurationRoot = builder.Build();

        var host = Host.CreateDefaultBuilder()
            .ConfigureServices((context, services) =>
            {
                services.AddSingleton<IAppConfig, AppConfig>(
                    _ => configurationRoot.GetSection(nameof(AppConfig)).Get<AppConfig>());
            })
            .Build();

        _appConfig = ActivatorUtilities.GetServiceOrCreateInstance<IAppConfig>(host.Services);

        IronXL.License.LicenseKey = _appConfig.LicenseKey;
        IronXLOld.License.LicenseKey = _appConfig.LicenseKey;

        Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.InvariantCulture;
    }
}