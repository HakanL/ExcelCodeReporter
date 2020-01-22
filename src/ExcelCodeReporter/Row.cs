using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Haukcode.ExcelCodeReporter
{
    public class Row
    {
        private ExcelWriter writer;
        private ExcelWorksheet backer;
        private int row;
        private int currentCol;
        private Action<ExcelStyle> defaultStyle;

        internal Row(ExcelWriter writer, ExcelWorksheet backer, int row, Action<ExcelStyle> defaultStyle)
        {
            this.writer = writer;
            this.backer = backer;
            this.row = row;
            this.defaultStyle = defaultStyle;
            this.currentCol = 0;
        }

        public int CurrentCol => this.currentCol;

        public Row AddHeader(object caption, double? width = null, Action<ExcelStyle> style = null)
        {
            this.currentCol++;

            this.backer.SetValue(this.row, this.currentCol, caption, style: s =>
                {
                    this.defaultStyle?.Invoke(s);
                    style?.Invoke(s);
                });

            if (width != null)
                this.backer.Column(this.currentCol).Width = width.Value;

            return this;
        }

        public Row Add(object value, string format = null, Action<ExcelStyle> style = null)
        {
            this.currentCol++;

            return SetAt(this.currentCol, value, format, style);
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

        public Row SetMaxFreezeColumn()
        {
            this.writer.MaxFreezeCol = this.currentCol;

            return this;
        }

        public Row SetMaxPrintAreaColumn()
        {
            this.writer.MaxPrintAreaCol = this.currentCol;

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
