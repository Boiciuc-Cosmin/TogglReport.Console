using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TogglReport.ConsoleApp.Dtos;

namespace TogglReport.ConsoleApp.Repository {
    public interface ITogglRepository {
        Task<DetailedReportDto> GetDetailsByMonth(string apiToken, int workspaceId, DateTime since, DateTime until);
        Task<List<WorkspaceDto>> GetWorkspaces(string apiToken);
    }
}