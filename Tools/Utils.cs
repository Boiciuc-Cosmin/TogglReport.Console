using System.Collections.Generic;
using System.Linq;
using TogglReport.ConsoleApp.Dtos;

namespace TogglReport.ConsoleApp.Tools {
    public static class Utils {

        public static List<IGrouping<string, ProjectData>> GetProjectsGroupedByDescription(DetailedReportDto detailedReport) {
            return detailedReport.Data.OrderBy(x => x.Start).GroupBy(x => x.Description).ToList();
        }
    }

}
