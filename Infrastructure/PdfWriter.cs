using SelectPdf;
using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TogglReport.ConsoleApp.Dtos;
using TogglReport.ConsoleApp.Tools;

namespace TogglReport.ConsoleApp.Infrastructure {
    public class PdfWriter {
        private const string PdfExtension = ".pdf";
        private readonly StringBuilder _stringBuilder;
        private const string DateTimeFormat = "MM.dd.yyyy_HH-mm";
        private readonly string DefaultFileName = $"PDF_Report({DateTime.Now.ToString(DateTimeFormat)}){PdfExtension}";
        private readonly string _filePathToSave;
        private readonly ILogger _logger;

        public PdfWriter(ILogger logger, string filePath) {
            if (Utils.HasFileNameWithExtension(filePath, PdfExtension)) {
                _filePathToSave = filePath;
            } else {
                _filePathToSave = Path.Combine(filePath, DefaultFileName);
            }

            _stringBuilder = new StringBuilder(@"<html>
                    <head>
                        <link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css'>
                        <script src='https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js'></script> 
                        <script src='https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/js/bootstrap.min.js'></script>
                        <script type='text/javascript' src='https://www.gstatic.com/charts/loader.js'></script>
                        <script type='text/javascript'>
                            google.charts.load('current', {'packages':['corechart']});
                        </script>
                        </head>");
            _logger = logger;
        }

        public async Task WritePdfFileAsync(DetailedReportDto detailedReport, GeneralProjectInformationDto generalInfo) {
            _logger.Information("PDF process started.");
            var stopWatch = new Stopwatch();
            stopWatch.Start();
            await Task.Run(() => {
                _stringBuilder.AppendLine("<body>");
                _stringBuilder.AppendLine("<div class='container'>");
                WriteGeneralInformation(generalInfo);
                GenerateBarGraphGroupedByDate(detailedReport.Data);
                GeneratePieGraphGroupedByProject(detailedReport.Data);
                WriteInformationGroupedByProject(detailedReport.Data);
                _stringBuilder.AppendLine("</div>");
                _stringBuilder.AppendLine("</body></html>");
            });

            HtmlToPdf converter = new HtmlToPdf();
            PdfDocument doc = converter.ConvertHtmlString(_stringBuilder.ToString());

            doc.Save(_filePathToSave);
            stopWatch.Stop();
            _logger.Information($"Pdf file created. Time for creation '{stopWatch.Elapsed.TotalSeconds.ToString("0.00")}' seconds");
            doc.Close();
        }

        private void WriteGeneralInformation(GeneralProjectInformationDto generalInfo) {
            _logger.Information("PDF: Writing general information.");
            var totalTimeSpan = TimeSpan.FromMilliseconds(generalInfo.TotalTime);
            _stringBuilder.AppendLine(@$"<div class='row'>
                                        <h2>Summary Report -> {generalInfo.User}</h2><br>
                                        <span style='text-align: left; color: grey; margin-top:0px; font-size:20;'>{generalInfo.SinceDateTime.ToString("MM/dd/yyyy")} - {generalInfo.UntilDateTime.ToString("MM/dd/yyyy")}</span><br>
                                        <b style='text-align: left; color: grey'; margin-top:3px;>TOTAL HOURS: {(int)totalTimeSpan.TotalHours}:{totalTimeSpan.Minutes}:{totalTimeSpan.Seconds}</b>
                                        </div>");
        }

        private void GeneratePieGraphGroupedByProject(List<ProjectData> projectDatas) {
            _logger.Information("PDF: Generating the graphs.");
            var projectsGroup = projectDatas.GroupBy(x => new { x.Project, Tag = String.Join(" ", x.Tags) }).ToList();
            _stringBuilder.AppendLine("<div class='row'><div id='piechart' ></div></div>");
            _stringBuilder.AppendLine(@"<script type='text/javascript'> google.charts.load('current', { 'packages': ['corechart'] });
                                    google.charts.setOnLoadCallback(drawChart);
                                        function drawChart() {
                                            var data = google.visualization.arrayToDataTable([");
            _stringBuilder.AppendLine("['Project', 'Total time'],");
            foreach (var project in projectsGroup) {
                long totalMiliseconds = project.Sum(x => x.Dur);
                var projectDuration = TimeSpan.FromMilliseconds(totalMiliseconds);
                _stringBuilder.AppendLine($"['{project.Key.Project} {project.Key.Tag} -> {(int)projectDuration.TotalHours}:{projectDuration.Minutes}:{projectDuration.Seconds}', {projectDuration.TotalHours}],");
            }
            _stringBuilder.AppendLine(@"]);
             var options = { 
                    legend: {alignment:'center', position:'right', maxLines: 1, textStyle:{fontSize: 15}},
                    chartArea: {left: 0, width: 600, height:250},   
                    height: 300, width: 600
               };

            var chart = new google.visualization.PieChart(document.getElementById('piechart'));
            chart.draw(data, options);
            }");
            _stringBuilder.AppendLine("</script>");
        }

        private void GenerateBarGraphGroupedByDate(List<ProjectData> projectDatas) {
            var projectsGroup = projectDatas.OrderBy(x => x.Start).GroupBy(x => x.Start.Value.ToString("ddd \\\\nMM/dd")).ToList();
            _stringBuilder.AppendLine("<div class='row' style='margin-left: -150px;'><div id='chart_div'>test</div></div>");
            _stringBuilder.AppendLine(@"<script type='text/javascript'> google.charts.load('current', { 'packages': ['bar'] });
                                    google.charts.setOnLoadCallback(drawChart);
                                        function drawChart() {
                                            var data = google.visualization.arrayToDataTable([");
            _stringBuilder.AppendLine("['DateTime', 'Total time'],");
            foreach (var day in projectsGroup) {
                long totalMiliseconds = day.Sum(x => x.Dur);
                var projectDuration = TimeSpan.FromMilliseconds(totalMiliseconds);
                _stringBuilder.AppendLine($"['{day.Key}', {projectDuration.TotalHours}],");
            }
            _stringBuilder.AppendLine(@"]);
             var options = { 'width': 1250, 'height': 300, legend: { position: 'none' },
                    vAxis: {
                    gridlines: {count: 9},
                    minValue: 0,
                    ticks: [1,2,3,4,5,6,7,8,9]
                  },
                    hAxis: {
                        gridlines:{ minSpacing:20},
                        showTextEvery:1,
                        slantedText: false,
                   }
               };

            var chart = new google.visualization.ColumnChart(document.getElementById('chart_div'));
            chart.draw(data, options);
            }");
            _stringBuilder.AppendLine("</script>");
        }

        private void WriteInformationGroupedByProject(List<ProjectData> projectDatas) {
            _logger.Information("PDF: Writing all data entries.");
            var projectsGroup = projectDatas.GroupBy(x => x.Project).ToList();
            _stringBuilder.AppendLine("<div class='row'> <div class='col-sm-10'><div style='color:grey; font-weight: normal; margin: 10px 0px 5px -12px'>PROJECT - TIME ENTRY</div></div><div class='col-sm-2'><div style='color:grey; font-weight: normal; text-align:right; margin: 10px 10px 5px 0px'>DURATION</div></div></div>");
            _stringBuilder.AppendLine("<div class='row'> <div class='panel-group'>");
            foreach (var project in projectsGroup) {
                long totalProjectMiliseconds = project.Sum(x => x.Dur);
                var projectDuration = TimeSpan.FromMilliseconds(totalProjectMiliseconds);
                _stringBuilder.AppendLine(@$"<div class='panel panel-primary'>
                                             <div class='panel-heading'><div class='row'><div class='col-sm-10'>{project.Key}</div><div class='col-sm-2'><div style='text-align:right; margin-right: 17px'>{(int)projectDuration.TotalHours}:{projectDuration.Minutes.ToString("D2")}:{projectDuration.Seconds.ToString("D2")}</div></div></div></div> ");
                var tagGrouped = project.GroupBy(x => new { Tag = String.Join(" ", x.Tags) });
                foreach (var tag in tagGrouped) {
                    long totalTagMiliseconds = tag.Sum(x => x.Dur);
                    var tagDuration = TimeSpan.FromMilliseconds(totalTagMiliseconds);
                    int descriptionTimeMargin = 32;
                    if (!string.IsNullOrEmpty(tag.Key.Tag)) {
                        _stringBuilder.AppendLine(@$"
                                            <div class='panel-content' style='margin-left: 30px; margin-top: 10px; margin-bottom: 10px; margin-right: 15px;'>
                                            <div class='panel panel-info'>
                                                <div class='panel-heading'><div class='row'><div class='col-sm-10'>{tag.Key.Tag}</div><div class='col-sm-2'><div style='text-align:right; margin-right: 1px'>{(int)tagDuration.TotalHours}:{tagDuration.Minutes.ToString("D2")}:{tagDuration.Seconds.ToString("D2")}</div></div></div></div>");
                        descriptionTimeMargin = 15;
                    }

                    var descriptionGrouped = tag.GroupBy(x => x.Description);
                    foreach (var description in descriptionGrouped) {
                        long totalDescriptionMiliseconds = description.Sum(x => x.Dur);
                        var descriptionDuration = TimeSpan.FromMilliseconds(totalDescriptionMiliseconds);
                        _stringBuilder.AppendLine(@$"<div class='panel-content' style='margin-top: 10px; margin-bottom: 10px;'>
                                                   <div class='row'> <div class='col-sm-10'><div style='text-align:left; margin-left: 20px'>{description.Key}</div></div> <div class='col-sm-2'><div style='text-align:right; margin-right: {descriptionTimeMargin}px'>{(int)descriptionDuration.TotalHours}:{descriptionDuration.Minutes.ToString("D2")}:{descriptionDuration.Seconds.ToString("D2")}</div></div></div>
                                                  </div>");
                    }

                    if (!string.IsNullOrEmpty(tag.Key.Tag)) {
                        _stringBuilder.AppendLine("</div></div>");
                    }
                }

                _stringBuilder.AppendLine("</div>");
            }

            _stringBuilder.AppendLine("</div></div>");
        }


    }
}
