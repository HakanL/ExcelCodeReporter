using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Haukcode.ExcelCodeReporter
{
    public class Row
    {
        private readonly ExcelWriter writer;
        private readonly ExcelWorksheet backer;
        private readonly int row;
        private Action<ExcelStyle> defaultStyle;

        internal Row(ExcelWriter writer, ExcelWorksheet backer, int row, Action<ExcelStyle> defaultStyle)
        {
            this.writer = writer;
            this.backer = backer;
            this.row = row;
            this.defaultStyle = defaultStyle;
            CurrentCol = 0;
        }

        public int CurrentCol { get; private set; }

        public Row AddHeader(object caption, double? width = null, Action<ExcelStyle> style = null)
        {
            CurrentCol++;

            this.backer.SetValue(this.row, CurrentCol, caption, style: s =>
            {
                this.defaultStyle?.Invoke(s);
                style?.Invoke(s);
            });

            if (width != null)
                this.backer.Column(CurrentCol).Width = width.Value;

            return this;
        }

        public Row Add(object value = null, string format = null, Action<ExcelStyle> style = null)
        {
            CurrentCol++;

            return SetAt(CurrentCol, value, format, style);
        }

        public Row SetAt(int col, object value, string format = null, Action<ExcelStyle> style = null)
        {
            this.backer.SetValue(this.row, col, value, style: s =>
            {
                this.defaultStyle?.Invoke(s);
                style?.Invoke(s);
                if (format != null)
                    s.Numberformat.Format = format;
            });

            return this;
        }

        public Row SetFormulaAt(int col, string formula, string format = null, Action<ExcelStyle> style = null)
        {
            this.backer.SetFormula(this.row, col, formula, style: s =>
            {
                this.defaultStyle?.Invoke(s);
                style?.Invoke(s);
                if (format != null)
                    s.Numberformat.Format = format;
            });

            return this;
        }

        public Row SetMaxFreezeColumn()
        {
            this.writer.MaxFreezeCol = CurrentCol;

            return this;
        }

        public Row SetMaxPrintAreaColumn()
        {
            this.writer.MaxPrintAreaCol = CurrentCol;

            return this;
        }

        public Row SetBorder(
            ExcelBorderStyle left = ExcelBorderStyle.None,
            ExcelBorderStyle top = ExcelBorderStyle.None,
            ExcelBorderStyle right = ExcelBorderStyle.None,
            ExcelBorderStyle bottom = ExcelBorderStyle.None)
        {
            var range = this.backer.Cells[$"{this.row}:{this.row}"];

            range.Style.Border.Bottom.Style = bottom;
            range.Style.Border.Top.Style = top;
            range.Style.Border.Left.Style = left;
            range.Style.Border.Right.Style = right;

            return this;
        }
    }
}
