using System;

namespace TogglReport.ConsoleApp.Dtos {
    public class GeneralProjectInformationDto {
        public GeneralProjectInformationDto(string user, long totalTime, int numberOfEntries, DateTime sinceDateTime, DateTime untilDateTime) {
            User = user;
            TotalTime = totalTime;
            NumberOfEntries = numberOfEntries;
            SinceDateTime = sinceDateTime;
            UntilDateTime = untilDateTime;
        }

        public string User { get; set; }
        public long TotalTime { get; set; }
        public int NumberOfEntries { get; set; }
        public DateTime SinceDateTime { get; set; }
        public DateTime UntilDateTime { get; set; }
    }
}
