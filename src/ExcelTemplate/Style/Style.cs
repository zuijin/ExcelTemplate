using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;

namespace ExcelTemplate.Style
{
    public class Style : IStyle
    {
        public bool ShrinkToFit { get; set; }
        public string DataFormat { get; set; }
        public IFont Font { get; set; }
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
            var obj = (Style)MemberwiseClone();
            obj.Font = new Font();
            obj.Font.CloneStyleFrom(this.Font);

            return obj;
        }

        ICellStyle _cellStyle;
        public ICellStyle GetCellStyle(IWorkbook workbook)
        {
            if (_cellStyle == null)
            {
                _cellStyle = StyleUtil.GetCellStyle(workbook, this);
            }

            return _cellStyle;
        }
    }
}
