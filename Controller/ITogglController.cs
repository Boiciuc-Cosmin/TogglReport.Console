using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TogglReport.ConsoleApp.Dtos;

namespace TogglReport.ConsoleApp.Controller {
    public interface ITogglController {
        Task<DetailedReportDto> GetDetailsByMonth(string apiToken, int workspaceId, DateTime since, DateTime until);
        Task<List<WorkspaceDto>> GetWorkspaces(string apiToken);
    }
}