using CommandLine;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Serilog;
using System;
using System.IO;
using System.Threading.Tasks;
using TogglReport.ConsoleApp.Controller;
using TogglReport.ConsoleApp.Dtos;
using TogglReport.ConsoleApp.Dtos.Options;
using TogglReport.ConsoleApp.Infrastructure;
using TogglReport.ConsoleApp.Repository;

namespace TogglReport.ConsoleApp {
    static class Program {
        static async Task Main(string[] args) {
            var builder = new ConfigurationBuilder();
            BuildConfig(builder);

            Log.Logger = new LoggerConfiguration()
                .ReadFrom.Configuration(builder.Build())
                .Enrich.FromLogContext()
                .WriteTo.Console()
                .CreateLogger();

            Log.Logger.Information("Application starting");

            var host = Host.CreateDefaultBuilder()
                .ConfigureServices((context, services) => {
                    var configurationRoot = context.Configuration;
                    services.AddHttpClient();
                    services.AddTransient<ITogglRepository, TogglRepository>();
                    services.AddTransient<ITogglController, TogglController>();
                    services.AddTransient<IFileWriterHandler, FileWriterHandler>();
                    services.AddTransient<IApplication, Application>();
                    services.AddSingleton(Log.Logger);
                    services.Configure<ApiOptions>(configurationRoot.GetSection("ConnectionStrings"));
                    services.Configure<ApiTokensOptions>(configurationRoot.GetSection("ApiTokens"));
                })
                .UseSerilog()
                .Build();

            var application = host.Services.GetService<IApplication>();

            await Parser.Default.ParseArguments<CommandOptions>(args)
                                .WithParsedAsync(async (opt) => await GetSelectedOptionsAsync(opt, application));

            Log.Logger.Information("Application closing");
        }

        static async Task GetSelectedOptionsAsync(CommandOptions options, IApplication application) {
            var argumentOptions = new ArgumentOptionsModel() {
                FilePath = options.FilePath,
                OutputType = options.OutputType,
                PeriodOfTime = options.PeriodOfTime,
                Since = options.Since,
                Until = options.Until,
                ApiToken = options.ApiToken,
                Workspace = options.Workspace
            };

            try {
                await application.RunAsync(argumentOptions);
            }
            catch (Exception ex) {
                Log.Logger.Error(ex.Message);
            }
        }

        static void BuildConfig(IConfigurationBuilder builder) {
            builder.SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
        }
    }
}
