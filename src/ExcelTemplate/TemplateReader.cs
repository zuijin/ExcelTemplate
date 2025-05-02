using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ExcelTemplate.Helper;
using ExcelTemplate.Model;
using NPOI.SS.UserModel;
using NPOI.Util;

namespace ExcelTemplate
{
    public class TemplateReader
    {
        IWorkbook _workbook;
        TemplateDesign _design;
        Type _type;

        List<CellException> _exceptionList = new List<CellException>();

        /// <summary>
        /// 
        /// </summary>
        public TemplateReader(IWorkbook workbook, Type type)
        {
            _workbook = workbook;
            _design = TypeDesignAnalysis.DesignAnalysis(type);
            _type = type;
        }

        /// <summary>
        /// 从文件中生成 ExcelWrap
        /// </summary>
        public static TemplateReader FromFile(Stream file, Type type)
        {
            var workbook = WorkbookFactory.Create(file);
            return new TemplateReader(workbook, type);
        }

        /// <summary>
        /// 是否存在转换错误
        /// </summary>
        /// <returns></returns>
        public bool IsError => _exceptionList.Any();

        /// <summary>
        /// 获取数据集合
        /// </summary>
        /// <returns></returns>
        public object GetData()
        {
            var obj = Read();
            if (IsError)
            {
                throw _exceptionList[0];
            }

            return obj;
        }

        /// <summary>
        /// 生成错误提示Excel
        /// </summary>
        /// <returns></returns>
        public IWorkbook BuildErrorExcel()
        {
            var newWorkbook = _workbook.Copy();
            var sheet = newWorkbook.GetSheetAt(0);
            var drawing = sheet.CreateDrawingPatriarch();
            var helper = _workbook.GetCreationHelper();

            foreach (var ex in _exceptionList)
            {
                var cell = sheet.GetRow(ex.Row).GetCell(ex.Column);
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
        public string BuildErrorMessage()
        {
            if (!IsError)
            {
                return string.Empty;
            }

            StringBuilder sb = new StringBuilder(500);
            foreach (var ex in _exceptionList.OrderBy(a => a.Row))
            {
                sb.AppendLine(ex.Message);
            }

            return sb.ToString();
        }

        /// <summary>
        /// 生成错误提示文件
        /// </summary>
        /// <returns></returns>
        public byte[] BuildErrorFile()
        {
            using (var ms = new MemoryStream())
            {
                var book = BuildErrorExcel();
                book.Write(ms);

                return ms.ToArray();
            }
        }

        /// <summary>
        /// 读取表单数据
        /// </summary>
        private object Read()
        {
            var data = Activator.CreateInstance(_type);
            var sheet = _workbook.GetSheetAt(0);
            var valueBlocks = _design.Where(a => a.BlockType == BlockType.Variable);
            var props = _type.GetProperties();

            foreach (var block in valueBlocks)
            {
                var row = sheet.GetRow(block.Position.Row);
                var cell = row.GetCell(block.Position.Col, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                var cellVal = GetCellValue(cell);

                ObjectHelper.SetObjectValue(data, block.ValuePath, cellVal);
            }

            var arrBlocks = _design.Where(a => a.BlockType == BlockType.Array);
            var groupBlocks = arrBlocks.GroupBy(a => a.ValuePath.Split('.')[0]);
            foreach (var group in groupBlocks)
            {
                var fieldName = group.Key;
                var prop = props.FirstOrDefault(a => a.Name == fieldName);
                if (prop != null)
                {
                    if (!TypeHelper.IsSubclassOfRawGeneric(typeof(List<>), prop.PropertyType))
                    {
                        throw new Exception("只支持 List<T> 类型的集合");
                    }

                    var list = ReadList(sheet, group.ToList(), prop.PropertyType);
                    prop.SetValue(data, list);
                }
            }

            return data;
        }

        /// <summary>
        /// 读取列表数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="blocks"></param>
        /// <param name="listType"></param>
        /// <returns></returns>
        private object ReadList(ISheet sheet, List<Block> blocks, Type listType)
        {
            object listObj = Activator.CreateInstance(listType);
            Type elementType = listType.GenericTypeArguments[0];
            var rowIndex = blocks.Max(a => Math.Max(a.Position.Row, a.MergeTo?.Row ?? 0)) + 1;

            while (true)
            {
                var row = sheet.GetRow(rowIndex);
                if (IsEmptyRow(row))
                {
                    break;
                }

                object obj = Activator.CreateInstance(elementType);
                foreach (var block in blocks)
                {
                    var cell = row.GetCell(block.Position.Col, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    var val = GetCellValue(cell);
                    if (val != null)
                    {
                        var fieldPath = block.ValuePath.Substring(block.ValuePath.IndexOf('.') + 1);
                        ObjectHelper.SetObjectValue(obj, fieldPath, val);
                    }
                }

                ObjectHelper.AddItemToList(listObj, obj);
                rowIndex++;
            }

            return listObj;
        }

        /// <summary>
        /// 判断是否空行
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private static bool IsEmptyRow(IRow row)
        {
            if (row == null)
            {
                return true;
            }

            var enumerator = row.GetEnumerator();
            while (enumerator.MoveNext())
            {
                if (enumerator.Current != null && !string.IsNullOrWhiteSpace(enumerator.Current.ToString()))
                {
                    return false;
                }
            }

            return true;
        }

        object GetCellValue(ICell cell)
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

            //return System.Convert.ChangeType(val, type);

            return val;
        }


        /// <summary>
        /// 判断是否可 null 类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        private bool IsNullableType(Type type)
        {
            return !type.IsValueType
                || type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>);
        }





        /// <summary>
        /// 添加错误提示
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="colName"></param>
        /// <param name="message"></param>
        /// <exception cref="NotImplementedException"></exception>
        public void AddError(string colName, string message)
        {
            //var col = _fieldInfos.Find(a => a.FieldName == colName || a.Title == colName);
            //int colIndex = col?.ColIndex ?? -1;
            //int rowIndexd = col?.RowIndex ?? -1;

            //_exceptionList.Add(new CellException(rowIndexd, colIndex, message));
        }
    }
}
