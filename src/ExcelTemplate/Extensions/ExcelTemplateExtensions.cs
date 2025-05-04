using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.Util;

namespace ExcelTemplate.Extensions
{
    public static class ExcelTemplateExtensions
    {
        /// <summary>
        /// 是否存在异常
        /// </summary>
        /// <returns></returns>
        public static bool HasError(this TemplateReader template)
        {
            return template.Exceptions.Any();
        }

        /// <summary>
        /// 生成错误提示Excel
        /// </summary>
        /// <returns></returns>
        public static IWorkbook BuildErrorExcel(this TemplateReader template)
        {
            var newWorkbook = template.WorkBook.Copy();
            var sheet = newWorkbook.GetSheetAt(0);
            var drawing = sheet.CreateDrawingPatriarch();
            var helper = template.WorkBook.GetCreationHelper();

            foreach (var ex in template.Exceptions)
            {
                var cell = sheet.GetRow(ex.Position.Row).GetCell(ex.Position.Col);
                var anchor = helper.CreateClientAnchor();
                var comment = drawing.CreateCellComment(anchor);
                comment.String = helper.CreateRichTextString(ex.Message);
                cell.CellComment = comment;
            }

            return newWorkbook;
        }

        /// <summary>
        /// 生成错误提示
        /// </summary>
        /// <returns></returns>
        public static string BuildErrorMessage(this TemplateReader template)
        {
            if (!template.Exceptions.Any())
            {
                return string.Empty;
            }

            StringBuilder sb = new StringBuilder(500);
            foreach (var ex in template.Exceptions.OrderBy(a => a.Position.Row))
            {
                sb.AppendLine(ex.Message);
            }

            return sb.ToString();
        }

        /// <summary>
        /// 生成错误提示文件
        /// </summary>
        /// <returns></returns>
        public static byte[] BuildErrorFile(this TemplateReader template)
        {
            using (var ms = new MemoryStream())
            {
                var book = template.BuildErrorExcel();
                book.Write(ms);

                return ms.ToArray();
            }
        }
    }
}
