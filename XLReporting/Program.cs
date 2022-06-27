// See https://aka.ms/new-console-template for more information

using System.Reflection;
using Microsoft.Extensions.Configuration;
using XLReporting.Configuration;
using IronXL;

var environment = Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT");
var builder = new ConfigurationBuilder()
    .AddJsonFile($"appsettings.json", true, true)
    .AddJsonFile($"appsettings.{environment}.json", true, true)
    .AddUserSecrets(Assembly.GetExecutingAssembly(), true)
    .AddEnvironmentVariables();

var configurationRoot = builder.Build();
var appConfig = configurationRoot.GetSection(nameof(AppConfig)).Get<AppConfig>();

License.LicenseKey = appConfig.LicenseKey;