using System;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.UserModel.Extensions;

namespace ExcelTemplate.Style
{
    public static class StyleUtil
    {
        public static IStyle ConvertStyle(ISimpleStyle simpleStyle)
        {
            var style = new Style()
            {
                WrapText = simpleStyle.WrapText,
                VerticalAlignment = (NPOI.SS.UserModel.VerticalAlignment)simpleStyle.VerticalAlignment,
                Alignment = (NPOI.SS.UserModel.HorizontalAlignment)simpleStyle.Alignment,
                ShrinkToFit = simpleStyle.ShrinkToFit,
                DataFormat = simpleStyle.DataFormat,
                FillForegroundColor = simpleStyle.ForegroundColor,
                FillBackgroundColor = simpleStyle.BackgroundColor,
            };

            style.Font = new Font()
            {
                IsBold = simpleStyle.IsBold,
                FontHeightInPoints = simpleStyle.FontHeightInPoints,
            };

            return style;
        }

        /// <summary>
        /// 获取 IStyle
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        public static IStyle ConvertStyle(IWorkbook workbook, ICellStyle cellStyle)
        {
            var style = new Style()
            {
                Alignment = cellStyle.Alignment,
                BorderBottom = cellStyle.BorderBottom,
                BorderDiagonal = cellStyle.BorderDiagonal,
                BorderDiagonalColor = GetColorHexString(workbook, cellStyle.BorderDiagonalColor),
                BorderDiagonalLineStyle = cellStyle.BorderDiagonalLineStyle,
                BorderLeft = cellStyle.BorderLeft,
                BorderRight = cellStyle.BorderRight,
                BorderTop = cellStyle.BorderTop,
                BottomBorderColor = GetColorHexString(workbook, cellStyle.BottomBorderColor),
                DataFormat = cellStyle.GetDataFormatString(),
                FillBackgroundColor = GetColorHexString(workbook, cellStyle.FillBackgroundColor),
                FillForegroundColor = GetColorHexString(workbook, cellStyle.FillForegroundColor),
                FillPattern = cellStyle.FillPattern,
                Font = cellStyle.GetFont(workbook),
                Indention = cellStyle.Indention,
                IsHidden = cellStyle.IsHidden,
                IsLocked = cellStyle.IsLocked,
                IsQuotePrefixed = cellStyle.IsQuotePrefixed,
                LeftBorderColor = GetColorHexString(workbook, cellStyle.LeftBorderColor),
                RightBorderColor = GetColorHexString(workbook, cellStyle.RightBorderColor),
                Rotation = cellStyle.Rotation,
                ShrinkToFit = cellStyle.ShrinkToFit,
                TopBorderColor = GetColorHexString(workbook, cellStyle.TopBorderColor),
                VerticalAlignment = cellStyle.VerticalAlignment,
                WrapText = cellStyle.WrapText,
            };

            return style;
        }

        /// <summary>
        /// 获取 ICellStyle
        /// </summary>
        /// <param name="style"></param>
        /// <param name="cellStyle"></param>
        public static ICellStyle GetCellStyle(IWorkbook workbook, IStyle style)
        {
            var cellStyle = workbook.CreateCellStyle();

            cellStyle.Alignment = style.Alignment;
            cellStyle.BorderBottom = style.BorderBottom;
            cellStyle.BorderDiagonal = style.BorderDiagonal;
            cellStyle.BorderLeft = style.BorderLeft;
            cellStyle.BorderRight = style.BorderRight;
            cellStyle.BorderDiagonalLineStyle = style.BorderDiagonalLineStyle;
            cellStyle.BorderTop = style.BorderTop;
            cellStyle.FillPattern = style.FillPattern;
            cellStyle.Indention = style.Indention;
            cellStyle.IsHidden = style.IsHidden;
            cellStyle.IsLocked = style.IsLocked;
            cellStyle.IsQuotePrefixed = style.IsQuotePrefixed;
            cellStyle.Rotation = style.Rotation;
            cellStyle.ShrinkToFit = style.ShrinkToFit;
            cellStyle.VerticalAlignment = style.VerticalAlignment;
            cellStyle.WrapText = style.WrapText;

            if (style.Font != null)
            {
                var font = workbook.CreateFont();
                font.CloneStyleFrom(style.Font);
                cellStyle.SetFont(font);
            }

            if (!string.IsNullOrWhiteSpace(style.DataFormat))
            {
                IDataFormat dataFormat = workbook.CreateDataFormat();
                cellStyle.DataFormat = dataFormat.GetFormat(style.DataFormat);
            }

            if (cellStyle is XSSFCellStyle xssfStye)
            {
                if (!string.IsNullOrWhiteSpace(style.FillForegroundColor))
                {
                    xssfStye.SetFillForegroundColor(GetXSSFColor(style.FillForegroundColor));
                }
                if (!string.IsNullOrWhiteSpace(style.FillBackgroundColor))
                {
                    xssfStye.SetFillBackgroundColor(GetXSSFColor(style.FillBackgroundColor));
                }
                if (!string.IsNullOrWhiteSpace(style.LeftBorderColor))
                {
                    xssfStye.SetLeftBorderColor(GetXSSFColor(style.LeftBorderColor));
                }
                if (!string.IsNullOrWhiteSpace(style.RightBorderColor))
                {
                    xssfStye.SetRightBorderColor(GetXSSFColor(style.RightBorderColor));
                }
                if (!string.IsNullOrWhiteSpace(style.TopBorderColor))
                {
                    xssfStye.SetTopBorderColor(GetXSSFColor(style.TopBorderColor));
                }
                if (!string.IsNullOrWhiteSpace(style.BottomBorderColor))
                {
                    xssfStye.SetBottomBorderColor(GetXSSFColor(style.BottomBorderColor));
                }
                if (!string.IsNullOrWhiteSpace(style.FillForegroundColor))
                {
                    xssfStye.SetDiagonalBorderColor(GetXSSFColor(style.BorderDiagonalColor));
                }
            }
            else
            {
                cellStyle.FillForegroundColor = GetHSSFColor(workbook, style.FillForegroundColor);
                cellStyle.FillBackgroundColor = GetHSSFColor(workbook, style.FillBackgroundColor);
                cellStyle.LeftBorderColor = GetHSSFColor(workbook, style.LeftBorderColor);
                cellStyle.RightBorderColor = GetHSSFColor(workbook, style.RightBorderColor);
                cellStyle.TopBorderColor = GetHSSFColor(workbook, style.TopBorderColor);
                cellStyle.BottomBorderColor = GetHSSFColor(workbook, style.BottomBorderColor);
                cellStyle.BorderDiagonalColor = GetHSSFColor(workbook, style.BorderDiagonalColor);
            }

            return cellStyle;
        }

        /// <summary>
        /// 获取颜色对象
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="hexColor"></param>
        /// <returns></returns>
        public static short GetHSSFColor(IWorkbook workbook, string hexColor)
        {
            if (string.IsNullOrWhiteSpace(hexColor))
            {
                return IndexedColors.Automatic.Index;
            }

            var (r, g, b) = ParseHexRgb(hexColor);
            var palette = ((HSSFWorkbook)workbook).GetCustomPalette();

            // 查找已有颜色或添加新颜色
            var color = palette.FindColor(r, g, b) ?? palette.AddColor(r, g, b);
            return color.Indexed;
        }
        

        public static XSSFColor GetXSSFColor(string hexColor)
        {
            var (r, g, b) = ParseHexRgb(hexColor);
            return new XSSFColor(new byte[] { r, g, b });
        }


        /// <summary>
        /// 获取颜色对象
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="hexColor"></param>
        /// <returns></returns>
        public static string GetColorHexString(IWorkbook workbook, short colorIndex)
        {
            if (workbook is XSSFWorkbook)
            {
                return IndexedColors.ValueOf(colorIndex).HexString;
            }
            else
            {
                // 获取自定义调色板
                HSSFPalette palette = ((HSSFWorkbook)workbook).GetCustomPalette();

                // 通过索引获取颜色
                HSSFColor color = palette.GetColor(colorIndex);

                // 如果找不到颜色，返回黑色作为默认值
                return color.GetHexString();
            }
        }

        /// <summary>
        /// 解析RGB颜色字符串
        /// </summary>
        /// <param name="hexColor"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        public static (byte R, byte G, byte B) ParseHexRgb(string hexColor)
        {
            hexColor = hexColor?.Trim() ?? throw new ArgumentNullException(nameof(hexColor));

            // 支持 #RGB、#RRGGBB、RGB、RRGGBB 格式
            if (hexColor.StartsWith("#"))
                hexColor = hexColor.Substring(1);

            if (hexColor.Length == 3) // #RGB 格式
            {
                return (
                    R: Convert.ToByte(new string(hexColor[0], 2), 16),
                    G: Convert.ToByte(new string(hexColor[1], 2), 16),
                    B: Convert.ToByte(new string(hexColor[2], 2), 16)
                );
            }
            else if (hexColor.Length == 6) // #RRGGBB 格式
            {
                return (
                    R: Convert.ToByte(hexColor.Substring(0, 2), 16),
                    G: Convert.ToByte(hexColor.Substring(2, 2), 16),
                    B: Convert.ToByte(hexColor.Substring(4, 2), 16)
                );
            }

            throw new ArgumentException("不支持的十六进制颜色格式");
        }
    }
}
