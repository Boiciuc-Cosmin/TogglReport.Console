using System.Threading.Tasks;
using TogglReport.ConsoleApp.Dtos;

namespace TogglReport.ConsoleApp.Infrastructure {
    public interface IFileWriterHandler {
        Task WriteToExcelFileAsync(DetailedReportDto detailedReport, GeneralProjectInformationDto generalInfo, string filePathToSave);
        Task WriteToPdfFileAsync(DetailedReportDto detailedReport, GeneralProjectInformationDto generalInfo, string filePathToSave);
    }
}