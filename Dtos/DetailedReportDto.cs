using System;
using System.Collections.Generic;

namespace TogglReport.ConsoleApp.Dtos {
    public class DetailedReportDto {
        public int? Total_Grand { get; set; }
        public int? Total_Count { get; set; }
        public List<ProjectData> Data { get; set; }

    }

    public class ProjectData {
        public long? Id { get; set; }
        public string User { get; set; }
        public string Description { get; set; }
        public DateTime? Start { get; set; }
        public DateTime? End { get; set; }
        public long Dur { get; set; }
        public string Project { get; set; }
        public string Client { get; set; }
        public List<string> Tags { get; set; }
    }
}
