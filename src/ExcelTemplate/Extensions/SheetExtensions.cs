using System;
using System.Collections.Generic;
using System.Text;
using ExcelTemplate.Model;
using NPOI.SS.UserModel;

namespace ExcelTemplate.Extensions
{
    public static class SheetExtensions
    {
        public static ICell GetCell(this ISheet sheet, Position position)
        {
            return sheet.GetRow(position.Row).GetCell(position.Col, MissingCellPolicy.CREATE_NULL_AS_BLANK);
        }
    }
}
