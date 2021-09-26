using CommandLine;

namespace TogglReport.ConsoleApp.Dtos.Options {
    public class CommandOptions {
        [Option('f', "filepath", Required = true, HelpText = "Set where to be saved \n Optional: After the path add a file name followed by .xlsx extension")]
        public string FilePath { get; set; }

        [Option('t', "timePeriod", Required = false, HelpText = "Select a time period (this_month, last_month)")]
        public string PeriodOfTime { get; set; }

        [Option('s', "since", Required = false, HelpText = "Set since date time MM/dd/yyy")]
        public string Since { get; set; }

        [Option('u', "until", Required = false, HelpText = "Set until date time MM/dd/yyy")]
        public string Until { get; set; }

        [Option('o', "outputType", Required = false, HelpText = "Select output type (excel, pdf, both)")]
        public string OutputType { get; set; }

        [Option('a', "apiToken", Required = true, HelpText = "Set api token (cosmin, marius) or insert api token")]
        public string ApiToken { get; set; }

        [Option('w', "workspace", Required = false, HelpText = "Set workspace -> Optional")]
        public string Workspace { get; set; }
    }
}
