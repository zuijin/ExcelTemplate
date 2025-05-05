using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelTemplate.Model;
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
        public static bool HasError(this TemplateCapture template)
        {
            return template.Exceptions.Any();
        }

        /// <summary>
        /// 生成错误提示Excel
        /// </summary>
        /// <returns></returns>
        public static IWorkbook BuildErrorExcel(this IWorkbook workbook, List<CellException> exceptions)
        {
            var newWorkbook = workbook.Copy();
            var sheet = newWorkbook.GetSheetAt(0);
            var drawing = sheet.CreateDrawingPatriarch();
            var helper = workbook.GetCreationHelper();

            foreach (var ex in exceptions)
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
        public static string BuildErrorMessage(this IList<CellException> exceptions)
        {
            if (!exceptions.Any())
            {
                return string.Empty;
            }

            StringBuilder sb = new StringBuilder(500);
            foreach (var ex in exceptions.OrderBy(a => a.Position.Row))
            {
                sb.AppendLine(ex.Message);
            }

            return sb.ToString();
        }

        /// <summary>
        /// 生成错误提示文件
        /// </summary>
        /// <returns></returns>
        public static byte[] BuildErrorFile(this IWorkbook workbook, List<CellException> exceptions)
        {
            using (var ms = new MemoryStream())
            {
                var book = workbook.BuildErrorExcel(exceptions);
                book.Write(ms);

                return ms.ToArray();
            }
        }
    }
}
