using System;
using System.Collections.Generic;
using System.Text;
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
        public NPOI.SS.UserModel.HorizontalAlignment Alignment { get; set; }
        public bool WrapText { get; set; }
        public NPOI.SS.UserModel.VerticalAlignment VerticalAlignment { get; set; }
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
