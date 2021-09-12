using Serilog;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using TogglReport.ConsoleApp.Controller;
using TogglReport.ConsoleApp.Dtos;
using TogglReport.ConsoleApp.Infrastructure;

namespace TogglReport.ConsoleApp {

    public class Application {
        private readonly ITogglController _togglController;
        private readonly IFileWriterHandler _fileWriterHandler;
        private readonly ILogger _logger;

        public Application(ITogglController togglController, IFileWriterHandler fileWriterHandler, ILogger logger) {
            _togglController = togglController;
            _fileWriterHandler = fileWriterHandler;
            _logger = logger;
        }

        public async Task RunAsync(string filePathToSaveExcel) {
            if (string.IsNullOrEmpty(filePathToSaveExcel)) {
                throw new ArgumentException($"'{nameof(filePathToSaveExcel)}' cannot be null or empty.", nameof(filePathToSaveExcel));
            }

            var timeInterval = GetTimeInterval();
            string filePath = GetOutputType();
            var listOfWorkspaces = await _togglController.GetWorkspaces("e1068110f8dc1c37b020175003a3eb50");
            var detailedReport = new DetailedReportDto();

            foreach (var workspace in listOfWorkspaces) {
                detailedReport = await _togglController.GetDetailsByMonth("e1068110f8dc1c37b020175003a3eb50", workspace.Id, timeInterval.Since, timeInterval.Until);
            }

            var generalInfo = new GeneralProjectInformationDto(detailedReport.Data.First().User,
                                                                     detailedReport.Total_Grand,
                                                                     detailedReport.Total_Count,
                                                                     timeInterval.Since,
                                                                     timeInterval.Until);
            try {
                if (Path.GetExtension(filePath).Equals(".xlsx")) {
                    await _fileWriterHandler.WriteToExcelFileAsync(detailedReport, generalInfo, filePath);
                }
            }

            catch (Exception ex) {
                _logger.Error(ex.Message);
            }
        }

        private DateTimeSelectorDto GetTimeInterval() {
            Console.WriteLine("Select a period of time.");
            Console.WriteLine("1. This month");
            Console.WriteLine("2. Last month");
            Console.WriteLine("3. Custom date interval");
            Console.Write("Pick a number: ");

            switch (Console.ReadLine()) {
                case "1":
                    var lastDayOfMonth = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month);
                    var since = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                    var until = new DateTime(DateTime.Now.Year, DateTime.Now.Month, lastDayOfMonth);
                    return new DateTimeSelectorDto() { Since = since, Until = until };
                case "2":
                    var sinceLastMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month - 1, 1);
                    var untilLastMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month - 1, DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month - 1));
                    return new DateTimeSelectorDto() { Since = sinceLastMonth, Until = untilLastMonth };
                case "3":
                    break;
                default:
                    Console.WriteLine("The option you selected is not in the list please pick a number from the list");
                    GetTimeInterval();
                    break;
            }

            return new DateTimeSelectorDto();
        }

        private string GetOutputType() {
            Console.Clear();
            Console.WriteLine("Choose an output.");
            Console.WriteLine("1. EXCEL");
            Console.WriteLine("2. PDF");
            Console.Write("Pick a number: ");

            switch (Console.ReadLine()) {
                case "1":
                    Console.Write("Insert file path: ");
                    string filePath = Console.ReadLine();
                    if (!IsGoodPath(filePath, ".xlsx")) {
                        _logger.Error("File path does not contain a proper excel extension");
                        Environment.Exit(-1);
                    }
                    return filePath;

                case "2":
                    break;

                default:
                    Console.WriteLine("The option you selected is not in the list please pick a number from the list");
                    GetOutputType();
                    break;
            }

            return string.Empty;
        }

        private bool IsGoodPath(string path, string extension) {
            return !string.IsNullOrEmpty(path) && Path.HasExtension(path) && Path.GetExtension(path).Equals(extension);
        }
    }
}
