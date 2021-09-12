using Microsoft.Extensions.Options;
using Newtonsoft.Json;
using Serilog;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using TogglReport.ConsoleApp.Dtos;

namespace TogglReport.ConsoleApp.Repository {
    public class TogglRepository : ITogglRepository {
        private readonly IHttpClientFactory _httpClientFactory;
        private readonly ILogger _logger;
        private readonly IOptionsMonitor<ApiOptions> _apiOptions;

        public TogglRepository(IHttpClientFactory httpClientFactory, ILogger logger, IOptionsMonitor<ApiOptions> apiOptions) {
            _httpClientFactory = httpClientFactory;
            _logger = logger;
            _apiOptions = apiOptions;
        }

        public async Task<List<WorkspaceDto>> GetWorkspaces(string apiToken) {
            var apiPath = "api/v8/workspaces";
            var request = new HttpRequestMessage(HttpMethod.Get, $"{_apiOptions.CurrentValue.ApiUrl}/{apiPath}");
            request.Headers.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(ASCIIEncoding.ASCII.GetBytes($"{apiToken}:api_token")));

            using (var client = _httpClientFactory.CreateClient()) {
                try {
                    var response = await client.SendAsync(request);
                    var responseAsString = await response.Content.ReadAsStringAsync();
                    return JsonConvert.DeserializeObject<List<WorkspaceDto>>(responseAsString);
                }
                catch (Exception ex) {
                    _logger.Error(ex, ex.Message);
                    throw;
                }
            }
        }

        public async Task<DetailedReportDto> GetDetailsByMonth(string apiToken, int workspaceId, DateTime since, DateTime until) {
            var apiPath = "reports/api/v2/details";
            var uriBuilder = new UriBuilder($"{_apiOptions.CurrentValue.ApiUrl}/{apiPath}");
            var query = HttpUtility.ParseQueryString(uriBuilder.Query);

            query["user_agent"] = "cosma";
            query["workspace_id"] = workspaceId.ToString();
            query["since"] = since.ToString("yyyy-MM-dd");
            query["until"] = until.ToString("yyyy-MM-dd");
            uriBuilder.Query = query.ToString();

            var request = new HttpRequestMessage(HttpMethod.Get, uriBuilder.ToString());
            request.Headers.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(ASCIIEncoding.ASCII.GetBytes($"{apiToken}:api_token")));

            using (var client = _httpClientFactory.CreateClient()) {
                var response = await client.SendAsync(request);
                var resp0onseAsString = await response.Content.ReadAsStringAsync();
                return JsonConvert.DeserializeObject<DetailedReportDto>(resp0onseAsString);
            }
        }
    }
}
