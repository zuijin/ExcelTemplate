using System;
using NPOI.SS.UserModel;

namespace ExcelTemplate.Style
{
    public class ETFont : ICloneable, IEquatable<ETFont>
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

        public override bool Equals(object? obj)
        {
            return Equals(obj as ETFont);
        }

        public bool Equals(ETFont? other)
        {
            if (other == null) return false;
            return FontName == other.FontName &&
                   FontHeight == other.FontHeight &&
                   IsItalic == other.IsItalic &&
                   IsStrikeout == other.IsStrikeout &&
                   Color == other.Color &&
                   TypeOffset == other.TypeOffset &&
                   Underline == other.Underline &&
                   Charset == other.Charset &&
                   Index == other.Index &&
                   IsBold == other.IsBold;
        }

        public override int GetHashCode()
        {
            var hash = new System.HashCode();
            hash.Add(FontName);
            hash.Add(FontHeight);
            hash.Add(IsItalic);
            hash.Add(IsStrikeout);
            hash.Add(Color);
            hash.Add(TypeOffset);
            hash.Add(Underline);
            hash.Add(Charset);
            hash.Add(Index);
            hash.Add(IsBold);
            return hash.ToHashCode();
        }

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
