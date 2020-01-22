using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Haukcode.ExcelCodeReporter
{
    public class ExcelWriter : IDisposable
    {
        private ExcelPackage excelPackage;
        private ExcelWorksheet backer;
        private Action<ExcelStyle> currentHeaderStyle;
        private int firstHeaderRow;
        private int lastHeaderRow;
        private string title;

        public int CurrentRow { get; set; }

        public ExcelWriter()
        {
            this.excelPackage = new ExcelPackage();
        }

        public ExcelWriter(Stream input, string title, int worksheetIndex = 1)
        {
            this.title = title;
            this.excelPackage = new ExcelPackage(input);

            this.backer = this.excelPackage.Workbook.Worksheets[worksheetIndex];
            this.firstHeaderRow = int.MaxValue;
            this.lastHeaderRow = 0;
            CurrentRow = 0;
        }

        public void UseWorksheet(int worksheetIndex)
        {
            this.backer = this.excelPackage.Workbook.Worksheets[worksheetIndex];
            this.firstHeaderRow = int.MaxValue;
            this.lastHeaderRow = 0;
            CurrentRow = 0;
        }

        public static Stream GetStreamFromTempFile(string tempFileName)
        {
            var ms = new MemoryStream(File.ReadAllBytes(tempFileName));

            File.Delete(tempFileName);

            return ms;
        }

        public string SaveCloseAndGetFileName()
        {
            string tempOutputFileName = Path.GetTempFileName();
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
            this.firstHeaderRow = int.MaxValue;
            this.lastHeaderRow = 0;
            CurrentRow = 0;

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
            if (row.HasValue)
                CurrentRow = row.Value;
            else
                CurrentRow++;

            var newRow = new Row(this, this.backer, CurrentRow, defaultStyle: s =>
            {
                this.currentHeaderStyle?.Invoke(s);
                style?.Invoke(s);
            });

            if (height != null)
                this.backer.Row(CurrentRow).Height = height.Value;

            if (CurrentRow < this.firstHeaderRow)
                this.firstHeaderRow = CurrentRow;

            if (CurrentRow > this.lastHeaderRow)
                this.lastHeaderRow = CurrentRow;

            return newRow;
        }

        public void SetHeaderRow(int row)
        {
            CurrentRow = row;

            if (row < this.firstHeaderRow)
                this.firstHeaderRow = row;

            if (row > this.lastHeaderRow)
                this.lastHeaderRow = row;
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
            if (row.HasValue)
                CurrentRow = row.Value;
            else
                CurrentRow++;

            var newRow = new Row(this, this.backer, CurrentRow, style);

            return newRow;
        }

        public Row AddRow(object value, Action<ExcelStyle> style = null, int? row = null)
        {
            if (row.HasValue)
                CurrentRow = row.Value;
            else
                CurrentRow++;

            var newRow = new Row(this, this.backer, CurrentRow, style);

            newRow.Add(value);

            return newRow;
        }

        public ExcelWriter SetPrintArea()
        {
            if (MaxPrintAreaCol > 0)
                return SetPrintArea(1, 1, CurrentRow, MaxPrintAreaCol);
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
            if (this.lastHeaderRow > 0)
                this.backer.View.FreezePanes(this.lastHeaderRow + 1, MaxFreezeCol + 1);

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
            if (this.firstHeaderRow <= this.lastHeaderRow)
                this.backer.PrinterSettings.RepeatRows = new ExcelAddress($"{this.firstHeaderRow}:{this.lastHeaderRow}");

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
}
