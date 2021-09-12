namespace TogglReport.ConsoleApp.Dtos {
    public static class ExcelTemplateHeaders {
        //General information headers
        public const string User = "User";
        public const string TotalTime = "Total time";
        public const string NumberOfEntries = "Number of entries";
        public const string SelectedPeriod = "Selected period";

        //Table data headers
        public const string Project = "Project";
        public const string Tag = "Tag";
        public const string Client = "Client";
        public const string Description = "Description";
        public const string StartDateTime = "Start date time";
        public const string EndDateTime = "End date time";
        public const string Duration = "Duration";

        //Table titles
        public const string DataTableTitle = "Detailed informations";
        public const string TableTitleByProject = "Total time by project";
        public const string TableTitleByDescription = "Total time by description";
    }
}
