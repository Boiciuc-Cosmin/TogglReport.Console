using GemBox.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using TogglReport.ConsoleApp.Dtos;
using TogglReport.ConsoleApp.Tools;
using Excel = Microsoft.Office.Interop.Excel;

namespace TogglReport.ConsoleApp.Infrastructure {
    public class ExcelWriter : IDisposable {
        private const string TimeFormat = "h:mm:ss";
        private readonly Excel.Application xlApp;
        private readonly Workbook xlWorkbook;
        private readonly _Worksheet worksheet;
        private readonly string _filePathToSave;
        private readonly ILogger _logger;
        private readonly ExcelHelper _excelHelper;

        public ExcelWriter(ILogger logger, string filePathToSave) {
            if (!IsGoodPath(filePathToSave)) {
                throw new ArgumentException($"'{nameof(filePathToSave)}' is not a valid excel file path", nameof(filePathToSave));
            }

            _filePathToSave = filePathToSave;
            xlApp = new Excel.Application {
                DisplayAlerts = false
            };
            xlWorkbook = xlApp.Workbooks.Add();
            worksheet = (_Worksheet)xlWorkbook.Sheets[1];
            _excelHelper = new ExcelHelper(worksheet, xlWorkbook);
            _logger = logger;
        }

        public async Task WriteToExcelFileAsync(DetailedReportDto detailedReport, GeneralProjectInformationDto generalInfo) {
            if (!IsGeneralInformationValid(generalInfo)) {
                throw new ArgumentException($"{nameof(generalInfo)} is not a valid model", nameof(generalInfo));
            }

            _logger.Information("Excel process started...");

            var xlRange = worksheet.UsedRange;

            await WriteHeadersAsync(xlRange);
            WriteGeneralInformation(generalInfo, xlRange);
            await WriteDataInTable(detailedReport, xlRange);
            int lastRow = CreateTableForTotalTimeByProject(detailedReport.Data, xlRange);
            await WriteDataByDescription(detailedReport, xlRange, lastRow + 3);

            worksheet.Columns.AutoFit();
            xlWorkbook.SaveAs2(_filePathToSave);
            _logger.Information("Excel file created");
        }

        private void WriteGeneralInformation(GeneralProjectInformationDto generalProjectInformation, Excel.Range xlRange) {
            if (!IsGeneralInformationValid(generalProjectInformation)) {
                throw new ArgumentException($"{nameof(generalProjectInformation)} is not a valid model", nameof(generalProjectInformation));
            }

            _logger.Information("Excel: Writing general information");
            var totalTimeSpan = TimeSpan.FromMilliseconds(generalProjectInformation.TotalTime);
            xlRange.Cells[1, 2].Value2 = generalProjectInformation.User;
            xlRange.Cells[2, 2].Value2 = $"{(int)totalTimeSpan.TotalHours}:{totalTimeSpan.Minutes}:{totalTimeSpan.Seconds}";
            xlRange.Cells[3, 2].Value2 = generalProjectInformation.NumberOfEntries;
            xlRange.Cells[4, 2].Value2 = $"{generalProjectInformation.SinceDateTime.ToShortDateString()} -> {generalProjectInformation.UntilDateTime.ToShortDateString()}";
        }

        private async Task WriteDataInTable(DetailedReportDto detailedReport, Excel.Range xlRange) {
            _logger.Information("Excel: Writing data to the table");
            await Task.Run(() => {
                int row = 8;
                var sortedList = detailedReport.Data.OrderBy(x => x.Start).ToList();
                foreach (var report in sortedList) {
                    var duration = TimeSpan.FromMilliseconds(report.Dur);
                    xlRange.Cells[row, 1].Value2 = report.Project;
                    xlRange.Cells[row, 2].Value2 = _excelHelper.ConvertTagsToString(report.Tags);
                    xlRange.Cells[row, 3].Value2 = report.Client;
                    xlRange.Cells[row, 4].Value2 = report.Description;
                    xlRange.Cells[row, 5].Value2 = report.Start.ToString();
                    xlRange.Cells[row, 6].Value2 = report.End.ToString();
                    xlRange.Cells[row, 7].Value2 = $"{(int)duration.TotalHours}:{duration.Minutes}:{duration.Seconds}";
                    xlRange.Cells[row, 7].NumberFormat = TimeFormat;
                    _excelHelper.SetStyleForTableData(xlRange, row, startColumn: 1, endColumn: 7);
                    row++;
                }
            });
        }

        private async Task WriteHeadersAsync(Excel.Range xlRange) {
            _logger.Information("Excel: Writing headers");
            await Task.Run(() => {
                xlRange.Cells[1, 1].Value2 = ExcelTemplateHeaders.User;
                xlRange.Cells[2, 1].Value2 = ExcelTemplateHeaders.TotalTime;
                xlRange.Cells[3, 1].Value2 = ExcelTemplateHeaders.NumberOfEntries;
                xlRange.Cells[4, 1].Value2 = ExcelTemplateHeaders.SelectedPeriod;
                worksheet.Range[xlRange.Cells[1, 1], xlRange.Cells[4, 2]].Style = xlWorkbook.Styles[BuiltInCellStyleName.Accent3Pct20];
                worksheet.Range[xlRange.Cells[1, 1], xlRange.Cells[4, 2]].Style.Font.Size = 12;
                worksheet.Range[xlRange.Cells[1, 1], xlRange.Cells[4, 2]].Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;

                xlRange.Cells[7, 1].Value2 = ExcelTemplateHeaders.Project;
                xlRange.Cells[7, 2].Value2 = ExcelTemplateHeaders.Tag;
                xlRange.Cells[7, 3].Value2 = ExcelTemplateHeaders.Client;
                xlRange.Cells[7, 4].Value2 = ExcelTemplateHeaders.Description;
                xlRange.Cells[7, 5].Value2 = ExcelTemplateHeaders.StartDateTime;
                xlRange.Cells[7, 6].Value2 = ExcelTemplateHeaders.EndDateTime;
                xlRange.Cells[7, 7].Value2 = ExcelTemplateHeaders.Duration;

                _excelHelper.SetStyleForHeaders(xlRange, startRow: 7, startColumn: 1, endColumn: 7);
                _excelHelper.SetTitleStyleForTableTitle(xlRange, ExcelTemplateHeaders.DataTableTitle, startRow: 6, startColumn: 1, endColumn: 7);
                _excelHelper.SetTitleStyleForTableTitle(xlRange, ExcelTemplateHeaders.TableTitleByProject, startRow: 6, startColumn: 9, endColumn: 12);


                xlRange.Cells[7, 9].Value2 = ExcelTemplateHeaders.Project;
                xlRange.Cells[7, 10].Value2 = ExcelTemplateHeaders.Tag;
                xlRange.Cells[7, 11].Value2 = ExcelTemplateHeaders.Client;
                xlRange.Cells[7, 12].Value2 = ExcelTemplateHeaders.TotalTime;
                _excelHelper.SetStyleForHeaders(xlRange, startRow: 7, startColumn: 9, endColumn: 12);
            });
        }

        private int CreateTableForTotalTimeByProject(List<ProjectData> projectDatas, Excel.Range xlRange) {
            var projectsGroup = projectDatas.GroupBy(x => new { x.Project, Tag = String.Join(" ", x.Tags) }).ToList();
            int row = 7;
            foreach (var project in projectsGroup) {
                row++;
                long totalMiliseconds = project.Sum(x => x.Dur);
                WriteTableForTotalTimeByProject(xlRange, project.Key.Project, project.Key.Tag, project.FirstOrDefault().Client, totalMiliseconds, row);
            }

            _excelHelper.CreatePieChart(720, 50, worksheet.Range[xlRange.Cells[7, 9], xlRange.Cells[row, 12]], Convert.ToString(xlRange.Cells[6, 9].Value2));
            return row;
        }

        private void WriteTableForTotalTimeByProject(Excel.Range xlRange, string projectName, string tag, string client, long totalMiliseconds, int row) {
            var duration = TimeSpan.FromMilliseconds(totalMiliseconds);
            var totalTime = $"{(int)duration.TotalHours}:{duration.Minutes}:{duration.Seconds}";

            xlRange.Cells[row, 9].Value2 = projectName;
            xlRange.Cells[row, 10].Value2 = string.IsNullOrEmpty(tag) ? "Without tag" : tag;
            xlRange.Cells[row, 11].Value2 = client;
            xlRange.Cells[row, 12].Value2 = totalTime;
            _excelHelper.SetStyleForTableData(xlRange, row, startColumn: 9, endColumn: 12);
        }

        private async Task WriteDataByDescription(DetailedReportDto detailedReport, Excel.Range xlRange, int startRow) {
            await Task.Run(() => {
                int titlePosition = startRow;
                _excelHelper.SetTitleStyleForTableTitle(xlRange, ExcelTemplateHeaders.TableTitleByDescription, startRow, startColumn: 9, endColumn: 13);
                startRow++;
                xlRange.Cells[startRow, 9].Value2 = ExcelTemplateHeaders.Description;
                xlRange.Cells[startRow, 10].Value2 = ExcelTemplateHeaders.Project;
                xlRange.Cells[startRow, 11].Value2 = ExcelTemplateHeaders.Tag;
                xlRange.Cells[startRow, 12].Value2 = ExcelTemplateHeaders.Client;
                xlRange.Cells[startRow, 13].Value2 = ExcelTemplateHeaders.TotalTime;

                var projectsGroup = detailedReport.Data.OrderBy(x => x.Start).GroupBy(x => x.Description).ToList();
                foreach (var report in projectsGroup) {
                    startRow++;
                    var totalMiliseconds = report.Sum(x => x.Dur);
                    var duration = TimeSpan.FromMilliseconds(totalMiliseconds);
                    xlRange.Cells[startRow, 9].Value2 = report.Key;
                    xlRange.Cells[startRow, 10].Value2 = report.FirstOrDefault().Project;
                    xlRange.Cells[startRow, 11].Value2 = _excelHelper.ConvertTagsToString(report.FirstOrDefault().Tags);
                    xlRange.Cells[startRow, 12].Value2 = report.FirstOrDefault().Client;
                    xlRange.Cells[startRow, 13].Value2 = $"{(int)duration.TotalHours}:{duration.Minutes}:{duration.Seconds}";
                    xlRange.Cells[startRow, 13].NumberFormat = TimeFormat;
                    _excelHelper.SetStyleForTableData(xlRange, startRow, startColumn: 9, endColumn: 13);
                }
                _excelHelper.CreatePieChart(720, 400, worksheet.Range[xlRange.Cells[titlePosition + 1, 9], xlRange.Cells[startRow, 13]], Convert.ToString(xlRange.Cells[titlePosition, 9].Value2));
            });
        }

        private bool IsGeneralInformationValid(GeneralProjectInformationDto generalProjectInformation) {
            return generalProjectInformation != null && !string.IsNullOrEmpty(generalProjectInformation.User);
        }

        private bool IsGoodPath(string path) {
            return !string.IsNullOrEmpty(path) && Path.HasExtension(path) && Path.GetExtension(path).Equals(".xlsx");
        }

        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing) {
            xlWorkbook.Close(false);
            Marshal.FinalReleaseComObject(xlWorkbook);

            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
        }
    }
}
