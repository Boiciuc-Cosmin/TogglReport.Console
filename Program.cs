using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using SelectPdf;
using Serilog;
using System;
using System.IO;
using System.Linq;
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
            await application.RunAsync();
        }

        static void BuildConfig(IConfigurationBuilder builder) {
            builder.SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
        }
    }

    public class Application {
        private readonly ITogglController _togglController;
        private readonly IFileWriterHandler _fileWriterHandler;
        private readonly ILogger _logger;

        public Application(ITogglController togglController, IFileWriterHandler fileWriterHandler, ILogger logger) {
            _togglController = togglController;
            _fileWriterHandler = fileWriterHandler;
            _logger = logger;
        }

        public async Task RunAsync() {
            var listOfWorkspaces = await _togglController.GetWorkspaces("e1068110f8dc1c37b020175003a3eb50");
            var detailedReport = new DetailedReportDto();
            var sinceDate = new DateTime(2021, 9, 1);
            var untilDate = DateTime.Now;
            foreach (var workspace in listOfWorkspaces) {
                detailedReport = await _togglController.GetDetailsByMonth("e1068110f8dc1c37b020175003a3eb50", workspace.Id, sinceDate, untilDate);
            }

            var generalInfo = new GeneralProjectInformationDto(detailedReport.Data.First().User,
                                                                     detailedReport.Total_Grand,
                                                                     detailedReport.Total_Count,
                                                                     sinceDate,
                                                                     untilDate);
            try {
                await _fileWriterHandler.WriteToExcelFileAsync(detailedReport, generalInfo, @"D:\Personal_Projects\ToyApps\TogglReport.Console\bin\Debug\net5.0\file.xlsx");
            }

            catch (Exception ex) {
                _logger.Error(ex.Message);
            }

            string text = @"<html>
                         <body>
                          Hello World from selectpdf.com.
                         </body>
                        </html>
                        ";

            HtmlToPdf converter = new HtmlToPdf();
            // create a new pdf document converting an url
            PdfDocument doc = converter.ConvertHtmlString(text);

            // save pdf document
            doc.Save("Sample.pdf");

            // close pdf document
            doc.Close();
        }
    }
}
