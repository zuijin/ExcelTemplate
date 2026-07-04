using NPOI.SS.UserModel;

namespace ExcelTemplate.Style
{
    public class ETStyle : IETStyle
    {
        public bool ShrinkToFit { get; set; }
        public string DataFormat { get; set; }
        public ETFont Font { get; set; }
        public bool IsHidden { get; set; }
        public bool IsLocked { get; set; }
        public bool IsQuotePrefixed { get; set; }
        public HorizontalAlignment Alignment { get; set; }
        public bool WrapText { get; set; }
        public VerticalAlignment VerticalAlignment { get; set; }
        public short Rotation { get; set; }
        public short Indention { get; set; }
        public BorderStyle BorderLeft { get; set; }
        public BorderStyle BorderRight { get; set; }
        public BorderStyle BorderTop { get; set; }
        public BorderStyle BorderBottom { get; set; }
        public string LeftBorderColor { get; set; }
        public string RightBorderColor { get; set; }
        public string TopBorderColor { get; set; }
        public string BottomBorderColor { get; set; }
        public FillPattern FillPattern { get; set; }
        public string FillBackgroundColor { get; set; }
        public string FillForegroundColor { get; set; }
        public string BorderDiagonalColor { get; set; }
        public BorderStyle BorderDiagonalLineStyle { get; set; }
        public BorderDiagonal BorderDiagonal { get; set; }

        public override bool Equals(object? obj)
        {
            return Equals(obj as IETStyle);
        }

        public bool Equals(IETStyle? other)
        {
            if (other == null) return false;
            
            return ShrinkToFit == other.ShrinkToFit &&
                   DataFormat == other.DataFormat &&
                   (Font == other.Font || (Font != null && Font.Equals(other.Font))) &&
                   IsHidden == other.IsHidden &&
                   IsLocked == other.IsLocked &&
                   IsQuotePrefixed == other.IsQuotePrefixed &&
                   Alignment == other.Alignment &&
                   WrapText == other.WrapText &&
                   VerticalAlignment == other.VerticalAlignment &&
                   Rotation == other.Rotation &&
                   Indention == other.Indention &&
                   BorderLeft == other.BorderLeft &&
                   BorderRight == other.BorderRight &&
                   BorderTop == other.BorderTop &&
                   BorderBottom == other.BorderBottom &&
                   LeftBorderColor == other.LeftBorderColor &&
                   RightBorderColor == other.RightBorderColor &&
                   TopBorderColor == other.TopBorderColor &&
                   BottomBorderColor == other.BottomBorderColor &&
                   FillPattern == other.FillPattern &&
                   FillBackgroundColor == other.FillBackgroundColor &&
                   FillForegroundColor == other.FillForegroundColor &&
                   BorderDiagonalColor == other.BorderDiagonalColor &&
                   BorderDiagonalLineStyle == other.BorderDiagonalLineStyle &&
                   BorderDiagonal == other.BorderDiagonal;
        }

        public override int GetHashCode()
        {
            var hash = new System.HashCode();
            hash.Add(ShrinkToFit);
            hash.Add(DataFormat);
            hash.Add(Font);
            hash.Add(IsHidden);
            hash.Add(IsLocked);
            hash.Add(IsQuotePrefixed);
            hash.Add(Alignment);
            hash.Add(WrapText);
            hash.Add(VerticalAlignment);
            hash.Add(Rotation);
            hash.Add(Indention);
            hash.Add(BorderLeft);
            hash.Add(BorderRight);
            hash.Add(BorderTop);
            hash.Add(BorderBottom);
            hash.Add(LeftBorderColor);
            hash.Add(RightBorderColor);
            hash.Add(TopBorderColor);
            hash.Add(BottomBorderColor);
            hash.Add(FillPattern);
            hash.Add(FillBackgroundColor);
            hash.Add(FillForegroundColor);
            hash.Add(BorderDiagonalColor);
            hash.Add(BorderDiagonalLineStyle);
            hash.Add(BorderDiagonal);
            return hash.ToHashCode();
        }

        public object Clone()
        {
            var obj = (ETStyle)MemberwiseClone();
            if (this.Font != null)
            {
                obj.Font = (ETFont)this.Font.Clone();
            }

            return obj;
        }

        ICellStyle _cellStyle;
        public ICellStyle GetCellStyle(IWorkbook workbook)
        {
            if (_cellStyle == null)
            {
                _cellStyle = ETStyleUtil.GetCellStyle(workbook, this);
            }

            return _cellStyle;
        }
    }
}
