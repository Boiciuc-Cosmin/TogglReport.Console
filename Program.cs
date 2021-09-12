using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Serilog;
using System;
using System.IO;
using System.Threading.Tasks;
using TogglReport.ConsoleApp.Controller;
using TogglReport.ConsoleApp.Dtos;
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
                    services.AddSingleton(Log.Logger);
                    services.Configure<ApiOptions>(configurationRoot.GetSection("ConnectionStrings"));
                })
                .UseSerilog()
                .Build();


            var togglController = host.Services.GetService<ITogglController>();
            var fileWriterHandler = host.Services.GetService<IFileWriterHandler>();

            var application = new Application(togglController, fileWriterHandler, Log.Logger);
            try {
                await application.RunAsync(@"D:\Personal_Projects\ToyApps\TogglReport.Console\bin\Debug\net5.0\file.xlsx");
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
