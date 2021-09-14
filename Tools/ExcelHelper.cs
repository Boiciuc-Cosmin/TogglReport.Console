using GemBox.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace TogglReport.ConsoleApp.Tools {
    public class ExcelHelper {
        private readonly Workbook _xlWorkbook;

        public ExcelHelper(Workbook xlWorkbook) {
            _xlWorkbook = xlWorkbook;
        }

        public void SetTitleStyleForTableTitle(Excel.Range xlRange, _Worksheet worksheet, string title, int startRow, int startColumn, int? endRow = null, int? endColumn = null) {
            if (string.IsNullOrEmpty(title)) {
                throw new ArgumentException($"'{nameof(title)}' cannot be null or empty.", nameof(title));
            }

            endRow = endRow == null ? startRow : endRow;
            endColumn = endColumn == null ? startColumn : endColumn;

            worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].Value = title;
            worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].Style = _xlWorkbook.Styles[BuiltInCellStyleName.Accent2Pct20];
            worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].Merge();
        }

        public void SetStyleForHeaders(Excel.Range xlRange, _Worksheet worksheet, int startRow, int startColumn, int? endRow = null, int? endColumn = null) {
            endRow = endRow == null ? startRow : endRow;
            endColumn = endColumn == null ? startColumn : endColumn;

            worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].Style = _xlWorkbook.Styles[BuiltInCellStyleName.Accent2Pct40];
            worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        public void SetStyleForTableData(Excel.Range xlRange, _Worksheet worksheet, int startRow, int startColumn, int? endRow = null, int? endColumn = null) {
            endRow = endRow == null ? startRow : endRow;
            endColumn = endColumn == null ? startColumn : endColumn;

            var style = _xlWorkbook.Styles[BuiltInCellStyleName.Currency0];
            worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].Borders.LineStyle = XlLineStyle.xlContinuous;
            worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].Borders.Weight = XlBorderWeight.xlThin;
            worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].Style = style;
            worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].HorizontalAlignment = XlHAlign.xlHAlignLeft;
        }

        public string ConvertTagsToString(List<string> tags) {
            var stringBuilder = new StringBuilder();
            for (int i = 0; i < tags.Count; i++) {
                stringBuilder.Append(tags[i]);
                if (tags.Count != i && tags.Count > 1) {
                    stringBuilder.Append(", ");
                }
            }

            return stringBuilder.ToString();
        }

        public void CreatePieChart(int leftMargin, int topMargin, Excel.Range range, _Worksheet worksheet, string chartTitle) {
            var chartOgjs = (ChartObjects)worksheet.ChartObjects();
            var chartObj = chartOgjs.Add(leftMargin, topMargin, 400, 300);
            Chart xlChart = chartObj.Chart;
            xlChart.ChartType = XlChartType.xlPie;
            xlChart.SetSourceData(range, Type.Missing);
            xlChart.ChartTitle.Text = chartTitle;
            xlChart.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        }
    }
}
