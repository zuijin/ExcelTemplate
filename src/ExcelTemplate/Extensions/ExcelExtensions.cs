﻿using System;
using System.IO;
using System.Linq;
using ExcelTemplate.Model;
using ExcelTemplate.Style;
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

        /// <summary>
        /// 获取单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="position"></param>
        /// <returns></returns>
        public static ICell GetCell(this ISheet sheet, Position position)
        {
            return sheet.GetCell(position.Row, position.Col);
        }

        /// <summary>
        /// 获取单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public static ICell GetCell(this ISheet sheet, int row, int col)
        {
            var irow = sheet.GetRow(row);
            if (irow == null)
            {
                return null;
            }

            return irow.GetCell(col, MissingCellPolicy.CREATE_NULL_AS_BLANK);
        }

        /// <summary>
        /// 获取行，如果该行不存在，则创建一个
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public static IRow GetOrCreateRow(this ISheet sheet, int rowIndex)
        {
            var row = sheet.GetRow(rowIndex);
            if (row == null)
            {
                row = sheet.CreateRow(rowIndex);
            }

            return row;
        }

        /// <summary>
        /// 获取单元格，如果该单元格不存在，则创建一个
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="position"></param>
        /// <returns></returns>
        public static ICell GetOrCreateCell(this ISheet sheet, Position position)
        {
            var row = sheet.GetOrCreateRow(position.Row);
            return row.GetOrCreateCell(position.Col);
        }

        /// <summary>
        /// 添加单元格合并
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="position"></param>
        /// <param name="mergeTo"></param>
        public static void AddMergedRegion(this ISheet sheet, Position position, Position mergeTo)
        {
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(position.Row, mergeTo.Row, position.Col, mergeTo.Col));
        }

        /// <summary>
        /// 获取单元格，如果该单元格不存在，则创建一个
        /// </summary>
        /// <param name="row"></param>
        /// <param name="colIndex"></param>
        /// <returns></returns>
        public static ICell GetOrCreateCell(this IRow row, int colIndex)
        {
            var cell = row.GetCell(colIndex);
            if (cell == null)
            {
                cell = row.CreateCell(colIndex);
            }

            return cell;
        }

        /// <summary>
        /// 获取单元格的值
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static object GetValue(this ICell cell)
        {
            if (cell == null)
            {
                return null;
            }

            object val = null;
            switch (cell.CellType)
            {
                case CellType.String:
                    val = cell.StringCellValue;
                    break;
                case CellType.Numeric:
                    val = DateUtil.IsCellDateFormatted(cell) ? (object)cell.DateCellValue : cell.NumericCellValue;
                    break;
                case CellType.Boolean:
                    val = cell.BooleanCellValue;
                    break;
                case CellType.Blank:
                case CellType.Formula:
                case CellType.Unknown:
                    break;
                default:
                    val = cell.ToString();
                    break;
            }

            return val;
        }

        /// <summary>
        /// 获取单元格的值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public static object GetValue(this ICell cell, Type type)
        {
            var val = cell.GetValue();
            if (type == typeof(DateTime) && !(val is DateTime))
            {
                val = cell.DateCellValue;
            }

            val = Convert.ChangeType(val, type);
            return val;
        }

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="val"></param>
        public static void SetValue(this ICell cell, object val)
        {
            if (val == null)
            {
                return;
            }

            var numericTypes = new[]
            {
                typeof(byte), typeof(float), typeof(int), typeof(long), typeof(double), typeof(decimal),
                typeof(short), typeof(sbyte), typeof(ushort), typeof(uint), typeof(ulong),
            };

            if (val is string)
            {
                cell.SetCellValue((string)val);
            }
            else if (val is DateTime)
            {
                var dateTime = (DateTime)val;
                cell.SetCellValue(dateTime);

                //if (!DateUtil.IsCellDateFormatted(cell))
                //{
                //    IDataFormat dataFormat = cell.Sheet.Workbook.CreateDataFormat();
                //    ICellStyle style = cell.Sheet.Workbook.CreateCellStyle();

                //    if (dateTime == dateTime.Date)
                //    {
                //        style.DataFormat = dataFormat.GetFormat("yyyy/m/d");
                //    }
                //    else
                //    {
                //        style.DataFormat = dataFormat.GetFormat("yyyy/m/d h:mm:ss");
                //    }

                //    cell.CellStyle = style;
                //}
            }
            else if (val is bool)
            {
                cell.SetCellValue((bool)val);
            }
            else if (numericTypes.Contains(val.GetType()))
            {
                cell.SetCellValue(double.Parse(val.ToString()));
            }
            else
            {
                cell.SetCellValue(val.ToString());
            }
        }

        /// <summary>
        /// 设置单元格样式
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="style"></param>
        public static void SetStyle(this ICell cell, IETStyle style)
        {
            if (style != null)
            {
                cell.CellStyle = style.GetCellStyle(cell.Sheet.Workbook);
            }
        }

        /// <summary>
        /// 判断是否空单元格
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static bool IsEmpty(this ICell cell)
        {
            if (cell == null)
            {
                return true;
            }

            var val = cell.GetValue();
            if (val == null || string.IsNullOrWhiteSpace(val.ToString()))
            {
                return true;
            }

            return false;
        }
    }
}
