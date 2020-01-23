using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;

namespace Haukcode.ExcelCodeReporter
{
    public static class ExcelExtensions
    {
        public static void SetValue(this ExcelWorksheet ws, int row, int col, object value, string format = null, Action<ExcelStyle> style = null)
        {
            using (var range = ws.Cells[row, col])
            {
                range.Value = value;
                if (format != null)
                    range.Style.Numberformat.Format = format;

                style?.Invoke(range.Style);
            }
        }
    }
}