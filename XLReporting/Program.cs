using IronXL;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Serilog;
using System.Reflection;
using XLReporting.Configuration;
using XLReporting.Reporting;

Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.InvariantCulture;

var environment = Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT");
var builder = new ConfigurationBuilder()
    .AddJsonFile($"appsettings.json", true, true)
    .AddJsonFile($"appsettings.{environment}.json", true, true)
    .AddUserSecrets(Assembly.GetExecutingAssembly(), true)
    .AddEnvironmentVariables();

var configurationRoot = builder.Build();

var host = Host.CreateDefaultBuilder()
    .ConfigureServices((context, services) =>
    {
        services.AddSingleton<IAppConfig, AppConfig>(
            _ => configurationRoot.GetSection(nameof(AppConfig)).Get<AppConfig>());
        services.AddTransient<IReportGenerator, ReportGenerator>();
    })
    .UseSerilog()
    .Build();

License.LicenseKey = ActivatorUtilities.GetServiceOrCreateInstance<IAppConfig>(host.Services).LicenseKey;

var reportGenerator = ActivatorUtilities.CreateInstance<ReportGenerator>(host.Services);
reportGenerator.GenerateReport();