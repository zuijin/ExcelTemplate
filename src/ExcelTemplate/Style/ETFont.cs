using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;

namespace ExcelTemplate.Style
{
    public class ETFont : ICloneable
    {
        private double _fontHeightRaw = 11;

        public string FontName { get; set; } = "Calibri";
        public double FontHeight { get => _fontHeightRaw * 20; set => _fontHeightRaw = value / 20; }
        public double FontHeightInPoints { get => _fontHeightRaw; set => _fontHeightRaw = value; }
        public bool IsItalic { get; set; } = false;
        public bool IsStrikeout { get; set; } = false;
        public string Color { get; set; }
        public FontSuperScript TypeOffset { get; set; }
        public FontUnderlineType Underline { get; set; }
        public short Charset { get; set; }
        public short Index { get; }
        public bool IsBold { get; set; } = false;

        public object Clone()
        {
            return MemberwiseClone() as ETFont;
        }

        //public void CloneStyleFrom(IFont src)
        //{
        //    this.FontName = src.FontName;
        //    this.FontHeight = src.FontHeight;
        //    this.FontHeightInPoints = src.FontHeightInPoints;
        //    this.IsItalic = src.IsItalic;
        //    this.IsStrikeout = src.IsStrikeout;
        //    this.Color = src.Color;
        //    this.TypeOffset = src.TypeOffset;
        //    this.Underline = src.Underline;
        //    this.Charset = src.Charset;
        //    this.IsBold = src.IsBold;
        //}
    }
}
