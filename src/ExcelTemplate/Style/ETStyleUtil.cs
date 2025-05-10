using System;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.UserModel.Extensions;

namespace ExcelTemplate.Style
{
    public static class ETStyleUtil
    {
        public static IETStyle ConvertStyle(IETSimpleStyle simpleStyle)
        {
            var style = new ETStyle()
            {
                WrapText = simpleStyle.WrapText,
                VerticalAlignment = (VerticalAlignment)simpleStyle.VerticalAlignment,
                Alignment = (HorizontalAlignment)simpleStyle.Alignment,
                ShrinkToFit = simpleStyle.ShrinkToFit,
                DataFormat = simpleStyle.DataFormat,
                FillForegroundColor = simpleStyle.BgColor,
                FillBackgroundColor = simpleStyle.BgColor,
                FillPattern = FillPattern.SolidForeground,
            };

            style.Font = new ETFont()
            {
                IsBold = simpleStyle.IsBold,
                FontHeightInPoints = simpleStyle.FontHeightInPoints,
                Color = simpleStyle.TextColor,
            };

            return style;
        }

        /// <summary>
        /// 获取 IStyle
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        public static IETStyle ConvertStyle(IWorkbook workbook, ICellStyle cellStyle)
        {
            var style = new ETStyle()
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
                Font = ConvertFont(workbook, cellStyle.GetFont(workbook)),
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

            if (workbook is XSSFWorkbook)
            {
                if (cellStyle.FillBackgroundColorColor != null)
                {
                    style.FillBackgroundColor = RgbToHex(cellStyle.FillBackgroundColorColor.RGB);
                }

                if (cellStyle.FillForegroundColorColor != null)
                {
                    style.FillForegroundColor = RgbToHex(cellStyle.FillForegroundColorColor.RGB);
                }
            }

            return style;
        }

        public static IFont ConvertFont(IWorkbook workbook, ETFont font)
        {
            var ifont = workbook.CreateFont();
            ifont.Charset = font.Charset;
            ifont.FontHeight = font.FontHeight;
            ifont.FontHeightInPoints = font.FontHeightInPoints;
            ifont.FontName = font.FontName;
            ifont.IsBold = font.IsBold;
            ifont.IsItalic = font.IsItalic;
            ifont.IsStrikeout = font.IsStrikeout;
            ifont.TypeOffset = font.TypeOffset;
            ifont.Underline = font.Underline;

            if (!string.IsNullOrWhiteSpace(font.Color))
            {
                if (ifont is XSSFFont xf)
                {
                    xf.SetColor(GetXSSFColor(font.Color));
                }
                else
                {
                    ifont.Color = GetHSSFColor(workbook, font.Color);
                }
            }

            return ifont;
        }

        public static ETFont ConvertFont(IWorkbook workbook, IFont ifont)
        {
            var font = new ETFont()
            {
                Charset = ifont.Charset,
                FontHeight = ifont.FontHeight,
                FontHeightInPoints = ifont.FontHeightInPoints,
                FontName = ifont.FontName,
                IsBold = ifont.IsBold,
                IsItalic = ifont.IsItalic,
                IsStrikeout = ifont.IsStrikeout,
                TypeOffset = ifont.TypeOffset,
                Underline = ifont.Underline,
            };

            if (ifont is XSSFFont xf)
            {
                var c = xf.GetXSSFColor();
                if (c != null)
                {
                    font.Color = RgbToHex(c.RGB);
                }
            }
            else
            {
                font.Color = RgbToHex(((HSSFFont)ifont).GetHSSFColor((HSSFWorkbook)workbook).RGB);
            }

            return font;
        }

        /// <summary>
        /// 获取 ICellStyle
        /// </summary>
        /// <param name="style"></param>
        /// <param name="cellStyle"></param>
        public static ICellStyle GetCellStyle(IWorkbook workbook, IETStyle style)
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
                var font = ConvertFont(workbook, style.Font);
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
                if (!string.IsNullOrWhiteSpace(style.BorderDiagonalColor))
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

            var (r, g, b) = HexToRgb(hexColor);
            var palette = ((HSSFWorkbook)workbook).GetCustomPalette();

            // 查找已有颜色或添加新颜色
            var color = palette.FindColor(r, g, b) ?? palette.AddColor(r, g, b);
            return color.Indexed;
        }


        public static XSSFColor GetXSSFColor(string hexColor)
        {
            var (r, g, b) = HexToRgb(hexColor);
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
            if (colorIndex == 0)
            {
                return RgbToHex(IndexedColors.Automatic.RGB);
            }

            if (workbook is XSSFWorkbook)
            {
                return RgbToHex(IndexedColors.ValueOf(colorIndex).RGB);
            }
            else
            {
                // 获取自定义调色板
                HSSFPalette palette = ((HSSFWorkbook)workbook).GetCustomPalette();
                // 通过索引获取颜色
                HSSFColor color = palette.GetColor(colorIndex);
                // 如果找不到颜色，返回黑色作为默认值
                return RgbToHex(color.RGB);
            }
        }

        /// <summary>
        /// 判断是否符合十六进制RGB格式
        /// </summary>
        /// <param name="hexColor"></param>
        /// <returns></returns>
        public static bool IsRgbHex(string hexColor)
        {
            return Regex.IsMatch(hexColor, "^([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$");
        }

        /// <summary>
        /// 解析RGB颜色字符串
        /// </summary>
        /// <param name="hexColor"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        public static (byte R, byte G, byte B) HexToRgb(string hexColor)
        {
            if (!IsRgbHex(hexColor))
            {
                throw new ArgumentNullException(nameof(hexColor));
            }

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

        /// <summary>
        /// RGB转十六进制
        /// </summary>
        /// <param name="hexColor"></param>
        /// <returns></returns>
        public static string RgbToHex(byte[] rgb)
        {
            return $"{rgb[0]:X2}{rgb[1]:X2}{rgb[2]:X2}";
        }
    }
}
