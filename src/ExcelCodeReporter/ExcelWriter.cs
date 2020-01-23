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
        private readonly Dictionary<int, WorksheetData> worksheetData;
        private Action<ExcelStyle> currentHeaderStyle;
        private ExcelPackage excelPackage;
        private string title;

        public ExcelWriter()
        {
            this.worksheetData = new Dictionary<int, WorksheetData>();
            this.excelPackage = new ExcelPackage();
        }

        private void SetWorksheetData(int worksheetIndex, int firstHeaderRow, int lastHeaderRow, int currentRow)
        {
            if (this.worksheetData.TryGetValue(worksheetIndex, out var data))
            {
                data.FirstHeaderRow = firstHeaderRow;
                data.LastHeaderRow = lastHeaderRow;
                data.CurrentRow = currentRow;
            }
            else
            {
                this.worksheetData.Add(worksheetIndex, new WorksheetData
                {
                    FirstHeaderRow = firstHeaderRow,
                    LastHeaderRow = lastHeaderRow,
                    CurrentRow = currentRow
                });
            }
        }

        public ExcelWriter(Stream input, string title, int worksheetIndex = 1)
        {
            this.worksheetData = new Dictionary<int, WorksheetData>();
            this.title = title;
            this.excelPackage = new ExcelPackage(input);

            Backer = this.excelPackage.Workbook.Worksheets[worksheetIndex];
            SetWorksheetData(worksheetIndex, int.MaxValue, 0, 0);
        }

        public ExcelWriter UseWorksheet(int worksheetIndex)
        {
            Backer = this.excelPackage.Workbook.Worksheets[worksheetIndex];
            return this;
        }

        public ExcelWriter UseWorksheet(string name)
        {
            Backer = this.excelPackage.Workbook.Worksheets.First(x => x.Name == name);
            return this;
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

        public ExcelWorksheet Backer { get; private set; }

        public ExcelWriter AddWorksheet(string name)
        {
            var newExcelWorksheet = this.excelPackage.Workbook.Worksheets.Add(name);
            Backer = newExcelWorksheet;
            SetWorksheetData(Backer.Index, int.MaxValue, 0, 0);

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
            var data = this.worksheetData[Backer.Index];

            if (row.HasValue)
                data.CurrentRow = row.Value;
            else
                data.CurrentRow++;

            var newRow = new Row(this, Backer, data.CurrentRow, defaultStyle: s =>
            {
                this.currentHeaderStyle?.Invoke(s);
                style?.Invoke(s);
            });

            if (height != null)
                Backer.Row(data.CurrentRow).Height = height.Value;

            if (data.CurrentRow < data.FirstHeaderRow)
                data.FirstHeaderRow = data.CurrentRow;

            if (data.CurrentRow > data.LastHeaderRow)
                data.LastHeaderRow = data.CurrentRow;

            return newRow;
        }

        public void SetHeaderRow(int row)
        {
            var data = this.worksheetData[Backer.Index];

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
            var data = this.worksheetData[Backer.Index];

            if (row.HasValue)
                data.CurrentRow = row.Value;
            else
                data.CurrentRow++;

            var newRow = new Row(this, Backer, data.CurrentRow, style);

            return newRow;
        }

        public Row AddRow(object value, Action<ExcelStyle> style = null, int? row = null)
        {
            var data = this.worksheetData[Backer.Index];

            if (row.HasValue)
                data.CurrentRow = row.Value;
            else
                data.CurrentRow++;

            var newRow = new Row(this, Backer, data.CurrentRow, style);

            newRow.Add(value);

            return newRow;
        }

        public ExcelWriter SetPrintArea()
        {
            var data = this.worksheetData[Backer.Index];

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
            Backer.PrinterSettings.PrintArea = Backer.Cells[fromRow, fromCol, toRow, toCol];

            return this;
        }

        public ExcelWriter SetFreezeHeader()
        {
            var data = this.worksheetData[Backer.Index];

            if (data.LastHeaderRow > 0)
                Backer.View.FreezePanes(data.LastHeaderRow + 1, MaxFreezeCol + 1);

            return this;
        }

        public ExcelWriter SetOrientation(eOrientation orientation)
        {
            Backer.PrinterSettings.Orientation = orientation;

            return this;
        }

        public ExcelWriter SetFitOnePage()
        {
            Backer.PrinterSettings.FitToPage = true;

            return this;
        }

        public ExcelWriter SetFitToWidth(int pages = 1)
        {
            Backer.PrinterSettings.FitToPage = true;
            Backer.PrinterSettings.FitToWidth = pages;
            Backer.PrinterSettings.FitToHeight = 0;

            return this;
        }

        public ExcelWriter PrintGridLines(bool value = true)
        {
            Backer.PrinterSettings.ShowGridLines = value;

            return this;
        }

        public ExcelWriter PrintHeaderOnEachPage()
        {
            var data = this.worksheetData[Backer.Index];

            if (data.FirstHeaderRow <= data.LastHeaderRow)
            {
                Backer.PrinterSettings.RepeatRows =
                    new ExcelAddress($"{data.FirstHeaderRow}:{data.LastHeaderRow}");
            }

            return this;
        }

        public ExcelWriter PrintPageNumberInFooter(string pageNumberString)
        {
            Backer.HeaderFooter.OddFooter.RightAlignedText = string.Format(pageNumberString, ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);

            return this;
        }

        public ExcelWriter PrintTitleInFooter()
        {
            Backer.HeaderFooter.OddFooter.LeftAlignedText = this.title.Replace("&", "&&");

            return this;
        }

        public ExcelWriter PrintCenteredTextInFooter(string value)
        {
            Backer.HeaderFooter.OddFooter.CenteredText = value;

            return this;
        }
    }
}
