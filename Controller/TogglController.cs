using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TogglReport.ConsoleApp.Dtos;
using TogglReport.ConsoleApp.Repository;

namespace TogglReport.ConsoleApp.Controller {
    public class TogglController : ITogglController {
        private readonly ITogglRepository _togglRepository;

        public TogglController(ITogglRepository togglRepository) {
            _togglRepository = togglRepository;
        }

        public async Task<List<WorkspaceDto>> GetWorkspaces(string apiToken) {
            return await _togglRepository.GetWorkspaces(apiToken);
        }

        public async Task<DetailedReportDto> GetDetailsByMonth(string apiToken, int workspaceId, DateTime since, DateTime until) {
            return await _togglRepository.GetDetailsByMonth(apiToken, workspaceId, since, until);
        }
    }
}
