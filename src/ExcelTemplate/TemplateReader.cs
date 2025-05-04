using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelTemplate.Helper;
using ExcelTemplate.Model;
using NPOI.SS.UserModel;

namespace ExcelTemplate
{
    public class TemplateReader
    {
        IWorkbook _workbook;
        TemplateDesign _design;
        Type _type;
        List<CellException> _exceptions = new List<CellException>();

        public IWorkbook WorkBook { get => _workbook; }
        public TemplateDesign Design { get => _design; }
        public Type Type { get => _type; }
        public List<CellException> Exceptions { get => _exceptions; }

        /// <summary>
        /// 
        /// </summary>
        public TemplateReader(IWorkbook workbook, Type type, TemplateDesign design)
        {
            _workbook = workbook;
            _design = design;
            _type = type;
        }

        /// <summary>
        /// 创建 TemplateReader
        /// </summary>
        public static TemplateReader Create(Stream file, Type type)
        {
            var workbook = WorkbookFactory.Create(file);
            var design = TypeDesignAnalysis.DesignAnalysis(type);
            return new TemplateReader(workbook, type, design);
        }

        /// <summary>
        /// 获取数据
        /// </summary>
        /// <returns></returns>
        public object GetData()
        {
            var obj = Read(_workbook, _type, _design);
            return obj;
        }

        /// <summary>
        /// 读取表单数据
        /// </summary>
        private object Read(IWorkbook workBook, Type type, TemplateDesign design)
        {
            var data = Activator.CreateInstance(type);
            var sheet = workBook.GetSheetAt(0);
            var currentPage = (BlockPage)design.BlockPage.Clone();

            while (currentPage != null)
            {
                var blocks = currentPage.RowBlocks;
                var props = type.GetProperties();

                foreach (var block in blocks.OfType<ValueBlock>())
                {
                    var row = sheet.GetRow(block.Position.Row);
                    var cell = row.GetCell(block.Position.Col, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    var cellVal = GetCellValue(cell);

                    try
                    {
                        ObjectHelper.SetObjectValue(data, block.FieldPath, cellVal);
                    }
                    catch (Exception ex)
                    {
                        _exceptions.Add(new CellException(cell.RowIndex, cell.ColumnIndex, ex.Message, ex));
                    }
                }

                foreach (var table in blocks.OfType<TableBlock>())
                {
                    var prop = props.FirstOrDefault(a => a.Name == table.TableName);
                    if (prop != null)
                    {
                        if (!TypeHelper.IsSubclassOfRawGeneric(typeof(List<>), prop.PropertyType))
                        {
                            throw new Exception("只支持 List<T> 类型的集合");
                        }

                        var list = ReadList(sheet, table.Body, prop.PropertyType);
                        prop.SetValue(data, list);
                    }
                }

                currentPage = currentPage.Next;
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
        private object ReadList(ISheet sheet, List<TableBodyBlock> blocks, Type listType)
        {
            object listObj = Activator.CreateInstance(listType);
            Type elementType = listType.GenericTypeArguments[0];
            var rowIndex = blocks.First().Position.Row;

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
                        var fieldPath = block.FieldPath.Substring(block.FieldPath.IndexOf('.') + 1);

                        try
                        {
                            ObjectHelper.SetObjectValue(obj, fieldPath, val);
                        }
                        catch (Exception ex)
                        {
                            _exceptions.Add(new CellException(cell.RowIndex, cell.ColumnIndex, ex.Message, ex));
                        }
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

        private int FindNext(BlockPage page)
        {
            throw new NotImplementedException();
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
    }
}
