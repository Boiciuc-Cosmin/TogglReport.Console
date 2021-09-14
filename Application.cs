using Microsoft.Extensions.Options;
using Serilog;
using System;
using System.Linq;
using System.Threading.Tasks;
using TogglReport.ConsoleApp.Controller;
using TogglReport.ConsoleApp.Dtos;
using TogglReport.ConsoleApp.Dtos.Options;
using TogglReport.ConsoleApp.Infrastructure;

namespace TogglReport.ConsoleApp {

    public class Application : IApplication {
        private const string ExcelSelection = "excel";
        private const string PdfSelection = "pdf";
        private readonly ITogglController _togglController;
        private readonly IFileWriterHandler _fileWriterHandler;
        private readonly ILogger _logger;
        private readonly ApiTokensOptions _optionsMonitor;

        public Application(ITogglController togglController, IFileWriterHandler fileWriterHandler, ILogger logger, IOptionsMonitor<ApiTokensOptions> optionsMonitor) {
            _togglController = togglController;
            _fileWriterHandler = fileWriterHandler;
            _logger = logger;
            _optionsMonitor = optionsMonitor.CurrentValue;
        }

        public async Task RunAsync(ArgumentOptionsModel argumentOptions) {
            if (argumentOptions is null) {
                throw new ArgumentNullException(nameof(argumentOptions));
            }

            var timeInterval = GetTimeInterval(argumentOptions.PeriodOfTime, argumentOptions.Since, argumentOptions.Until);
            var apiToken = GetApiToken(argumentOptions.ApiToken);
            var workspaceId = await SelectWorkspaceId(apiToken);

            var detailedReport = await _togglController.GetDetailsByMonth(apiToken, workspaceId, timeInterval.Since, timeInterval.Until);

            if (detailedReport.Data.Count == 0 && detailedReport.Total_Grand is null) {
                _logger.Information("There is no data in this workspace");
                return;
            }

            var generalInfo = new GeneralProjectInformationDto(detailedReport.Data.First().User,
                                                                     (int)detailedReport.Total_Grand,
                                                                     (int)detailedReport.Total_Count,
                                                                     timeInterval.Since,
                                                                     timeInterval.Until);
            try {
                if (argumentOptions.OutputType.Equals(ExcelSelection, StringComparison.CurrentCultureIgnoreCase)) {
                    await _fileWriterHandler.WriteToExcelFileAsync(detailedReport, generalInfo, argumentOptions.FilePath);
                }
                if (argumentOptions.OutputType.Equals(PdfSelection, StringComparison.CurrentCultureIgnoreCase)) {
                    await _fileWriterHandler.WriteToPdfFileAsync();
                } else {
                    _logger.Information("We couldn't get the output type. Default output is Excel.");
                    await _fileWriterHandler.WriteToExcelFileAsync(detailedReport, generalInfo, argumentOptions.FilePath);
                }
            }

            catch (Exception ex) {
                _logger.Error(ex.Message);
            }
        }

        private DateTimeSelectorDto GetTimeInterval(string periodOfTime, string sinceDate, string untilDate) {
            if (!string.IsNullOrEmpty(periodOfTime)) {
                switch (periodOfTime) {
                    case "this_month":
                        var lastDayOfMonth = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month);
                        var since = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                        var until = new DateTime(DateTime.Now.Year, DateTime.Now.Month, lastDayOfMonth);
                        return new DateTimeSelectorDto() { Since = since, Until = until };
                    case "last_month":
                        var sinceLastMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddMonths(-1);
                        var untilLastMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month - 1, DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month - 1));
                        return new DateTimeSelectorDto() { Since = sinceLastMonth, Until = untilLastMonth };
                }
            }

            if (DateTime.TryParse(sinceDate, out DateTime sinceParsedDateTime) && DateTime.TryParse(untilDate, out DateTime untilParsedDateTime)) {
                return new DateTimeSelectorDto() { Since = sinceParsedDateTime, Until = untilParsedDateTime };
            } else {
                throw new ArgumentException("since/until date time are not in a correct format");
            }
        }

        private string GetApiToken(string ApiToken) {
            switch (ApiToken) {
                case "cosmin":
                    return _optionsMonitor.CosminToken;
                case "marius":
                    return _optionsMonitor.MariusToken;
                default:
                    return ApiToken;
            }
        }

        private async Task<int> SelectWorkspaceId(string apiToken) {
            var listOfWorkspaces = await _togglController.GetWorkspaces(apiToken);
            Console.WriteLine("Select workspace from the list");

            foreach (var workspace in listOfWorkspaces) {
                Console.WriteLine($"\t-> {workspace.Name}");
            }

            Console.Write("Enter name of workspace: ");
            var input = Console.ReadLine();
            var selectedWorkspace = listOfWorkspaces.FirstOrDefault(x => x.Name.Equals(input, StringComparison.CurrentCultureIgnoreCase));

            return selectedWorkspace != null && !string.IsNullOrEmpty(selectedWorkspace.Name)
                ? selectedWorkspace.Id
                : throw new ArgumentException("Selected workspace does not exist");
        }
    }
}
