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
        /// <summary>
        /// 自适应伸缩区块位置时，最大查找行数
        /// </summary>
        const int FIND_MAX_ROW = 100;

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

            DesignInspector.Check(design);
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
            var current = (BlockSection)design.BlockSection.Clone();
            int nextRow = 0;

            while (current != null)
            {
                // 重新匹配区块位置，原因是，如果上一个区块是 Table 的话，
                // 受到 Table 的数据影响，本身会占用更多的行，导致下一个区块的实际位置会跟模版定义时的位置不符
                ReMatchingPosition(current, sheet, nextRow);
                var lastRow = 0;

                if (DesignInspector.IsTableSection(current))
                {
                    lastRow = ReadTable(current, sheet, ref data);
                }
                else
                {
                    lastRow = ReadForm(current, sheet, ref data);
                }

                current = current.Next;
                nextRow = lastRow + 1;
            }

            return data;
        }

        /// <summary>
        /// 读取表单数据
        /// </summary>
        /// <param name="section"></param>
        /// <param name="sheet"></param>
        /// <param name="data"></param>
        /// <returns>返回最后一行的下标</returns>
        private int ReadForm(BlockSection section, ISheet sheet, ref object data)
        {
            var blocks = section.Blocks;
            var props = data.GetType().GetProperties();

            foreach (var block in blocks.OfType<ValueBlock>())
            {
                var row = sheet.GetRow(block.Position.Row);
                var cell = row.GetCell(block.Position.Col, MissingCellPolicy.CREATE_NULL_AS_BLANK);

                try
                {
                    var cellVal = GetCellValue(cell);
                    ObjectHelper.SetObjectValue(data, block.FieldPath, cellVal);
                }
                catch (Exception ex)
                {
                    _exceptions.Add(new CellException(cell.RowIndex, cell.ColumnIndex, ex.Message, ex));
                }
            }

            return blocks.Max(a => a.Position.Row);
        }

        /// <summary>
        /// 读取列表数据
        /// </summary>
        /// <param name="section"></param>
        /// <param name="sheet"></param>
        /// <param name="data"></param>
        /// <returns>返回最后一行的下标</returns>
        /// <exception cref="Exception"></exception>
        private int ReadTable(BlockSection section, ISheet sheet, ref object data)
        {
            var tables = section.Blocks.OfType<TableBlock>();
            var props = data.GetType().GetProperties();
            var lastRow = section.BeginRow;

            foreach (var table in tables)
            {
                var prop = props.FirstOrDefault(a => a.Name == table.TableName);
                if (prop != null)
                {
                    if (!TypeHelper.IsSubclassOfRawGeneric(typeof(List<>), prop.PropertyType))
                    {
                        throw new Exception("只支持 List<T> 类型的集合");
                    }

                    int itemCount;
                    var list = ReadOneList(sheet, table.Body, prop.PropertyType, out itemCount);
                    var tableLastRow = table.Header.Max(a => a.MergeTo?.Row ?? a.Position.Row) + itemCount; //列表最后一行 = 表头最后一行 + 表体数据行数
                    lastRow = Math.Max(lastRow, tableLastRow);

                    prop.SetValue(data, list);
                }
            }

            return lastRow;
        }

        /// <summary>
        /// 读取列表数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="blocks"></param>
        /// <param name="listType"></param>
        /// <param name="itemCount">数据行数</param>
        /// <returns></returns>
        private object ReadOneList(ISheet sheet, List<TableBodyBlock> blocks, Type listType, out int itemCount)
        {
            object listObj = Activator.CreateInstance(listType);
            Type elementType = listType.GenericTypeArguments[0];
            var rowIndex = blocks.First().Position.Row;
            var beginCol = blocks.Min(a => a.Position.Col);
            var endCol = blocks.Max(a => a.Position.Col);

            itemCount = 0;

            while (true)
            {
                var row = sheet.GetRow(rowIndex);
                if (IsEmptyRow(row, beginCol, endCol))
                {
                    break;
                }

                object obj = Activator.CreateInstance(elementType);
                foreach (var block in blocks)
                {
                    var cell = row.GetCell(block.Position.Col, MissingCellPolicy.CREATE_NULL_AS_BLANK);

                    try
                    {
                        var val = GetCellValue(cell);
                        if (val != null)
                        {
                            var fieldPath = block.FieldPath.Substring(block.FieldPath.IndexOf('.') + 1);

                            ObjectHelper.SetObjectValue(obj, fieldPath, val);
                        }
                    }
                    catch (Exception ex)
                    {
                        _exceptions.Add(new CellException(cell.RowIndex, cell.ColumnIndex, ex.Message, ex));
                    }
                }

                ObjectHelper.AddItemToList(listObj, obj);
                rowIndex++;
                itemCount++;
            }

            return listObj;
        }

        /// <summary>
        /// 判断是否空行
        /// </summary>
        /// <param name="row"></param>
        /// <param name="beginCol">限定起始列</param>
        /// <param name="endCol">限定结束列</param>
        /// <returns></returns>
        private static bool IsEmptyRow(IRow row, int beginCol = 0, int endCol = 0)
        {
            if (row == null)
            {
                return true;
            }

            var enumerator = row.GetEnumerator();
            while (enumerator.MoveNext())
            {
                var current = enumerator.Current;
                if (current.ColumnIndex < beginCol)
                {
                    continue;
                }

                if (endCol > 0 && current.ColumnIndex > endCol)
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(enumerator.Current.ToString()))
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// 获取单元格的值
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
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
        /// 重新匹配区块位置
        /// </summary>
        /// <param name="section"></param>
        /// <param name="sheet"></param>
        /// <param name="beginRow"></param>
        /// <exception cref="Exception"></exception>
        private void ReMatchingPosition(BlockSection section, ISheet sheet, int beginRow)
        {
            if (section == null)
            {
                return;
            }

            // 重新设定起始行
            var currentRow = beginRow;
            var minRow = section.BeginRow;
            if (minRow < currentRow)
            {
                section.ApplyOffset(currentRow - minRow, 0);
            }
            else
            {
                currentRow = minRow;
            }

            // 开始查找，跳过空行
            int offsetRow = 0;
            while (true)
            {
                var row = sheet.GetRow(currentRow);
                if (!IsEmptyRow(row))
                {
                    section.ApplyOffset(offsetRow, 0);
                    break;
                }

                currentRow++;
                offsetRow++;
                if (offsetRow >= FIND_MAX_ROW)
                {
                    throw new Exception($"已查找达到{FIND_MAX_ROW}行，未找到下一个模版区块");
                }
            }
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
