using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Haukcode.ExcelCodeReporter
{
    public static class Helper
    {
        public const string FormatMoney = "$###,###,##0.00";
        public const string FormatAccounting = @"_(""$""* #,##0.00_);_(""$""* \(#,##0.00\);_(""$""* ""-""??_);_(@_)";
        public const string FormatDate = "M/d/yyyy";
        public const string FormatPercent1 = "#0.0%";

        public static void ApplyDefaultStyling(ExcelWriter ws, string worksheetName, string reportTitle)
        {
            ws.AddWorksheet(worksheetName)
                .SetHeaderStyle(s =>
                {
                    s.Font.UnderLine = true;
                    s.Font.Bold = true;
                    s.Fill.PatternType = ExcelFillStyle.Solid;
                    s.Fill.BackgroundColor.SetColor(Color.Blue);
                    s.Font.Color.SetColor(Color.White);
                });

            ws.SetTitle(reportTitle, s => { s.Font.Size = 24; s.Font.Bold = true; });
        }

        public static void ApplyDefaultReportSettings(ExcelWriter ws, DateTime generatedTimestamp, eOrientation orientation)
        {
            ws.Backer.Calculate();

            ws.SetPrintArea()
                .SetOrientation(orientation)
                .SetFitToWidth()
                .SetFreezeHeader()
                .PrintGridLines()
                .PrintHeaderOnEachPage()
                .PrintTitleInFooter()
                .PrintPageNumberInFooter("Page {0} of {1}")
                .PrintCenteredTextInFooter(generatedTimestamp.Kind == DateTimeKind.Utc ? $"Generated {generatedTimestamp} UTC" : $"Generated {generatedTimestamp}");
        }
    }
}
