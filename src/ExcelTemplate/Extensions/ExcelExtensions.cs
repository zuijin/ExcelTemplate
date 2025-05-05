using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelTemplate.Model;
using NPOI.SS.UserModel;

namespace ExcelTemplate.Extensions
{
    public static class ExcelExtensions
    {
        /// <summary>
        /// 写入文件
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="filePath"></param>
        public static void Save(this IWorkbook workbook, string filePath)
        {
            var dir = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            using (var fs = File.OpenWrite(filePath))
            {
                workbook.Write(fs);
                fs.Close();
            }
        }

        public static ICell GetCell(this ISheet sheet, Position position)
        {
            return sheet.GetCell(position.Row, position.Col);
        }

        public static ICell GetCell(this ISheet sheet, int row, int col)
        {
            return sheet.GetRow(row).GetCell(col, MissingCellPolicy.CREATE_NULL_AS_BLANK);
        }

        public static IRow GetOrCreateRow(this ISheet sheet, int rowIndex)
        {
            var row = sheet.GetRow(rowIndex);
            if (row == null)
            {
                row = sheet.CreateRow(rowIndex);
            }

            return row;
        }

        public static ICell GetOrCreateCell(this ISheet sheet, Position position)
        {
            var row = sheet.GetOrCreateRow(position.Row);
            return row.GetOrCreateCell(position.Col);
        }

        public static void AddMergedRegion(this ISheet sheet, Position position, Position mergeTo)
        {
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(position.Row, mergeTo.Row, position.Col, mergeTo.Col));
        }

        public static ICell GetOrCreateCell(this IRow row, int colIndex)
        {
            var cell = row.GetCell(colIndex);
            if (cell == null)
            {
                cell = row.CreateCell(colIndex);
            }

            return cell;
        }

        public static object GetValue(this ICell cell)
        {
            object val = null;
            switch (cell.CellType)
            {
                case CellType.String:
                    val = cell.StringCellValue;
                    break;
                case CellType.Numeric:
                    val = DateUtil.IsCellDateFormatted(cell) ? (object?)cell.DateCellValue : cell.NumericCellValue;
                    break;
                case CellType.Boolean:
                    val = cell.BooleanCellValue;
                    break;
                case CellType.Blank:
                    break;
                default:
                    val = cell.ToString();
                    break;
            }

            return val;
        }
    }
}
