using Serilog;
using System;
using System.Threading.Tasks;
using TogglReport.ConsoleApp.Dtos;

namespace TogglReport.ConsoleApp.Infrastructure {
    public class FileWriterHandler : IFileWriterHandler {
        private readonly ILogger _logger;

        public FileWriterHandler(ILogger logger) {
            _logger = logger;
        }

        public async Task WriteToExcelFileAsync(DetailedReportDto detailedReport, GeneralProjectInformationDto generalInfo, string filePathToSave) {
            using (var excelWriter = new ExcelWriter(_logger, filePathToSave)) {
                await excelWriter.WriteToExcelFileAsync(detailedReport, generalInfo);
            }
        }

        public async Task WriteToPdfFileAsync() {
            throw new NotImplementedException();
        }
    }
}
