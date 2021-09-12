using GemBox.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace TogglReport.ConsoleApp.Tools {
    public class ExcelHelper {
        private readonly _Worksheet _worksheet;
        private readonly Workbook _xlWorkbook;

        public ExcelHelper(_Worksheet worksheet, Workbook xlWorkbook) {
            _worksheet = worksheet;
            _xlWorkbook = xlWorkbook;
        }

        public void SetTitleStyleForTableTitle(Excel.Range xlRange, string title, int startRow, int startColumn, int? endRow = null, int? endColumn = null) {
            if (string.IsNullOrEmpty(title)) {
                throw new ArgumentException($"'{nameof(title)}' cannot be null or empty.", nameof(title));
            }

            endRow = endRow == null ? startRow : endRow;
            endColumn = endColumn == null ? startColumn : endColumn;

            _worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].Value = title;
            _worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            _worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].Style = _xlWorkbook.Styles[BuiltInCellStyleName.Accent2Pct20];
            _worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].Merge();
        }

        public void SetStyleForHeaders(Excel.Range xlRange, int startRow, int startColumn, int? endRow = null, int? endColumn = null) {
            endRow = endRow == null ? startRow : endRow;
            endColumn = endColumn == null ? startColumn : endColumn;

            _worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].Style = _xlWorkbook.Styles[BuiltInCellStyleName.Accent2Pct40];
            _worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].Style.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        public void SetStyleForTableData(Excel.Range xlRange, int startRow, int startColumn, int? endRow = null, int? endColumn = null) {
            endRow = endRow == null ? startRow : endRow;
            endColumn = endColumn == null ? startColumn : endColumn;

            var style = _xlWorkbook.Styles[BuiltInCellStyleName.Currency0];
            _worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].Borders.LineStyle = XlLineStyle.xlContinuous;
            _worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].Borders.Weight = XlBorderWeight.xlThin;
            _worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].Style = style;
            _worksheet.Range[xlRange.Cells[startRow, startColumn], xlRange.Cells[endRow, endColumn]].HorizontalAlignment = XlHAlign.xlHAlignLeft;
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

        public void CreatePieChart(int leftMargin, int topMargin, Excel.Range range, string chartTitle) {
            var chartOgjs = (ChartObjects)_worksheet.ChartObjects();
            var chartObj = chartOgjs.Add(leftMargin, topMargin, 400, 300);
            Chart xlChart = chartObj.Chart;
            xlChart.ChartType = XlChartType.xlPie;
            xlChart.SetSourceData(range, Type.Missing);
            xlChart.ChartTitle.Text = chartTitle;
            xlChart.ApplyDataLabels(XlDataLabelsType.xlDataLabelsShowValue, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        }
    }
}
