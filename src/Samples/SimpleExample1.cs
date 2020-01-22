using System;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Haukcode.ExcelCodeReporter.Samples
{
    public class SimpleExample1
    {
        public string Execute()
        {
            using var ws = new ExcelWriter();

            Helper.ApplyDefaultStyling(ws, "Transactions", $"Charges for Sample Place");

            ws.AddRow($"Since: {DateTime.Today.AddMonths(-1):yyyy-MM-dd}");

            ws.AddRow();

            var headerRow = ws.AddHeaderRow()
                .AddHeader("Customer Id", 12)
                .AddHeader("First name", 30)
                .AddHeader("Last name", 30)
                .AddHeader("DOB", 30)
                .AddHeader("Trans. Date", 15)
                .AddHeader("Description", 90)
                .AddHeader("Amount", 15);

            var row = ws.AddRow()
                .Add(123456, "@")
                .Add("Firstname", "@")
                .Add("Lastname", "@")
                .Add(new DateTime(1976, 8, 22), Helper.FormatDate, style: s => s.HorizontalAlignment = ExcelHorizontalAlignment.Left)
                .Add(DateTime.Today, Helper.FormatDate, style: s => s.HorizontalAlignment = ExcelHorizontalAlignment.Left)
                .Add("Sample Charge", "@")
                .Add(123.45, Helper.FormatMoney);

            Helper.ApplyDefaultReportSettings(ws, DateTime.Now, eOrientation.Landscape);

            string filename = ws.SaveCloseAndGetFileName();

            return filename;
        }
    }
}
