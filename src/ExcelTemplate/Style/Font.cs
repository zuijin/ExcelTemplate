using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;

namespace ExcelTemplate.Style
{
    public class Font : IFont
    {
        private double _fontHeightRaw = 11;

        public string FontName { get; set; } = "Calibri";
        public double FontHeight { get => _fontHeightRaw * 20; set => _fontHeightRaw = value / 20; }
        public double FontHeightInPoints { get => _fontHeightRaw; set => _fontHeightRaw = value; }
        public bool IsItalic { get; set; } = false;
        public bool IsStrikeout { get; set; } = false;
        public short Color { get; set; } = 0;
        public FontSuperScript TypeOffset { get; set; }
        public FontUnderlineType Underline { get; set; }
        public short Charset { get; set; }
        public short Index { get; }
        public short Boldweight { get; set; }
        public bool IsBold { get; set; } = false;

        public void CloneStyleFrom(IFont src)
        {
            this.FontName = src.FontName;
            this.FontHeight = src.FontHeight;
            this.FontHeightInPoints = src.FontHeightInPoints;
            this.IsItalic = src.IsItalic;
            this.IsStrikeout = src.IsStrikeout;
            this.Color = src.Color;
            this.TypeOffset = src.TypeOffset;
            this.Underline = src.Underline;
            this.Charset = src.Charset;
            this.IsBold = src.IsBold;
        }
    }
}
