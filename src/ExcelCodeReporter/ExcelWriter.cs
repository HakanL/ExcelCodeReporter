using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Haukcode.ExcelCodeReporter
{
    public class ExcelWriter : IDisposable
    {
        private ExcelPackage excelPackage;
        private ExcelWorksheet backer;
        private Action<ExcelStyle> currentHeaderStyle;
        private string title;

        private Dictionary<int, WorksheetData> WorksheetData { get; set; }

        public ExcelWriter()
        {
            this.WorksheetData = new Dictionary<int, WorksheetData>();
            this.excelPackage = new ExcelPackage();
        }

        private void SetWorksheetData(int worksheetIndex, int firstHeaderRow, int lastHeaderRow, int currentRow)
        {
            if (this.WorksheetData.TryGetValue(worksheetIndex, out var data))
            {
                data.FirstHeaderRow = firstHeaderRow;
                data.LastHeaderRow = lastHeaderRow;
                data.CurrentRow = currentRow;
            }
            else
            {
                this.WorksheetData.Add(worksheetIndex, new WorksheetData
                {
                    FirstHeaderRow = firstHeaderRow,
                    LastHeaderRow = lastHeaderRow,
                    CurrentRow = currentRow
                });
            }
        }

        public ExcelWriter(Stream input, string title, int worksheetIndex = 1)
        {
            this.WorksheetData = new Dictionary<int, WorksheetData>();
            this.title = title;
            this.excelPackage = new ExcelPackage(input);

            this.backer = this.excelPackage.Workbook.Worksheets[worksheetIndex];
            this.SetWorksheetData(worksheetIndex, int.MaxValue, 0, 0);
        }

        public void UseWorksheet(int worksheetIndex)
        {
            this.backer = this.excelPackage.Workbook.Worksheets[worksheetIndex];
        }

        public void UseWorksheet(string name)
        {
            this.backer = this.excelPackage.Workbook.Worksheets.First(x => x.Name == name);
        }

        public static Stream GetStreamFromTempFile(string tempFileName)
        {
            var ms = new MemoryStream(File.ReadAllBytes(tempFileName));

            File.Delete(tempFileName);

            return ms;
        }

        public string SaveCloseAndGetFileName()
        {
            var tempOutputFileName = Path.GetTempFileName();
            File.Delete(tempOutputFileName);
            tempOutputFileName += ".xlsx";

            this.excelPackage.SaveAs(new FileInfo(tempOutputFileName));

            this.excelPackage.Dispose();
            this.excelPackage = null;

            return tempOutputFileName;
        }

        public ExcelWorksheet Backer => this.backer;

        public ExcelWriter AddWorksheet(string name)
        {
            var newExcelWorksheet = this.excelPackage.Workbook.Worksheets.Add(name);
            this.backer = newExcelWorksheet;
            this.SetWorksheetData(backer.Index, int.MaxValue, 0, 0);

            return this;
        }

        public void Dispose()
        {
            this.excelPackage?.Dispose();
            this.excelPackage = null;
        }

        public int MaxFreezeCol { get; set; }

        public int MaxPrintAreaCol { get; set; }

        public ExcelWriter SetHeaderStyle(Action<ExcelStyle> style)
        {
            this.currentHeaderStyle = style;

            return this;
        }

        public Row AddHeaderRow(double? height = null, Action<ExcelStyle> style = null, int? row = null)
        {
            var data = WorksheetData[this.Backer.Index];

            if (row.HasValue)
                data.CurrentRow = row.Value;
            else
                data.CurrentRow++;

            var newRow = new Row(this, this.backer, data.CurrentRow, defaultStyle: s =>
            {
                this.currentHeaderStyle?.Invoke(s);
                style?.Invoke(s);
            });

            if (height != null)
                this.backer.Row(data.CurrentRow).Height = height.Value;

            if (data.CurrentRow < data.FirstHeaderRow)
                data.FirstHeaderRow = data.CurrentRow;

            if (data.CurrentRow > data.LastHeaderRow)
                data.LastHeaderRow = data.CurrentRow;

            return newRow;
        }

        public void SetHeaderRow(int row)
        {
            var data = WorksheetData[this.Backer.Index];

            data.CurrentRow = row;

            if (row < data.FirstHeaderRow)
                data.FirstHeaderRow = row;

            if (row > data.LastHeaderRow)
                data.LastHeaderRow = row;
        }

        public Row SetTitle(string title, Action<ExcelStyle> style = null)
        {
            var row = AddRow();
            row.Add(title, style: style);

            this.title = title;

            return row;
        }

        public Row AddRow(Action<ExcelStyle> style = null, int? row = null)
        {
            var data = WorksheetData[this.Backer.Index];

            if (row.HasValue)
                data.CurrentRow = row.Value;
            else
                data.CurrentRow++;

            var newRow = new Row(this, this.backer, data.CurrentRow, style);

            return newRow;
        }

        public Row AddRow(object value, Action<ExcelStyle> style = null, int? row = null)
        {
            var data = WorksheetData[this.Backer.Index];

            if (row.HasValue)
                data.CurrentRow = row.Value;
            else
                data.CurrentRow++;

            var newRow = new Row(this, this.backer, data.CurrentRow, style);

            newRow.Add(value);

            return newRow;
        }

        public ExcelWriter SetPrintArea()
        {
            var data = WorksheetData[this.Backer.Index];

            if (MaxPrintAreaCol > 0)
                return SetPrintArea(1, 1, data.CurrentRow, MaxPrintAreaCol);
            else
                return this;
        }

        public ExcelWriter SetPrintArea(int toRow, int toCol)
        {
            return SetPrintArea(1, 1, toRow, toCol);
        }

        public ExcelWriter SetPrintArea(int fromRow, int fromCol, int toRow, int toCol)
        {
            this.backer.PrinterSettings.PrintArea = this.backer.Cells[fromRow, fromCol, toRow, toCol];

            return this;
        }

        public ExcelWriter SetFreezeHeader()
        {
            var data = WorksheetData[this.Backer.Index];

            if (data.LastHeaderRow > 0)
                this.backer.View.FreezePanes(data.LastHeaderRow + 1, MaxFreezeCol + 1);

            return this;
        }

        public ExcelWriter SetOrientation(eOrientation orientation)
        {
            this.backer.PrinterSettings.Orientation = orientation;

            return this;
        }

        public ExcelWriter SetFitOnePage()
        {
            this.backer.PrinterSettings.FitToPage = true;

            return this;
        }

        public ExcelWriter SetFitToWidth(int pages = 1)
        {
            this.backer.PrinterSettings.FitToPage = true;
            this.backer.PrinterSettings.FitToWidth = pages;
            this.backer.PrinterSettings.FitToHeight = 0;

            return this;
        }

        public ExcelWriter PrintGridLines(bool value = true)
        {
            this.backer.PrinterSettings.ShowGridLines = value;

            return this;
        }

        public ExcelWriter PrintHeaderOnEachPage()
        {
            var data = WorksheetData[this.Backer.Index];

            if (data.FirstHeaderRow <= data.LastHeaderRow)
                this.backer.PrinterSettings.RepeatRows =
                    new ExcelAddress($"{data.FirstHeaderRow}:{data.LastHeaderRow}");

            return this;
        }

        public ExcelWriter PrintPageNumberInFooter(string pageNumberString)
        {
            this.backer.HeaderFooter.OddFooter.RightAlignedText = string.Format(pageNumberString, ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);

            return this;
        }

        public ExcelWriter PrintTitleInFooter()
        {
            this.backer.HeaderFooter.OddFooter.LeftAlignedText = this.title.Replace("&", "&&");

            return this;
        }

        public ExcelWriter PrintCenteredTextInFooter(string value)
        {
            this.backer.HeaderFooter.OddFooter.CenteredText = value;

            return this;
        }
    }

    public class WorksheetData
    {
        public int FirstHeaderRow { get; set; }
        public int LastHeaderRow { get; set; }
        public int CurrentRow { get; set; }
    }
}
