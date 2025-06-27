using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.AccessControl;
using ExcelTemplate.Exceptions;
using ExcelTemplate.Extensions;
using ExcelTemplate.Helper;
using ExcelTemplate.Hint;
using ExcelTemplate.Model;
using NPOI.SS.UserModel;
using NPOI.Util;

namespace ExcelTemplate
{
    /// <summary>
    /// 负责按照模版定义从Excel中抓取数据
    /// </summary>
    public class TemplateCapture
    {
        /// <summary>
        /// 自适应伸缩区块位置时，最大查找行数
        /// </summary>
        const int FIND_MAX_ROW = 100;

        TemplateDesign _design;
        List<CellException> _exceptions;
        Dictionary<string, Func<object, object>> _dicMappings;

        public TemplateDesign Design { get => _design; }
        public List<CellException> Exceptions { get => _exceptions; }

        /// <summary>
        /// 
        /// </summary>
        public TemplateCapture(TemplateDesign design)
        {
            _design = design;
            _exceptions = new List<CellException>();
            _dicMappings = new Dictionary<string, Func<object, object>>();

            DesignInspector.Check(design);
        }

        /// <summary>
        /// 创建
        /// </summary>
        public static TemplateCapture Create(Type type)
        {
            var design = new TypeDesignAnalysis().DesignAnalysis(type);
            return new TemplateCapture(design);
        }

        public static TemplateCapture Create(string excelFile)
        {
            var design = new ExcelDesignAnalysis().DesignAnalysis(excelFile);
            return new TemplateCapture(design);
        }

        public T Capture<T>(Stream stream)
        {
            var workbook = WorkbookFactory.Create(stream);
            return Capture<T>(workbook);
        }

        public T Capture<T>(IWorkbook workbook)
        {
            var obj = (T)Capture(workbook, typeof(T));
            return obj;
        }

        public object Capture(Stream stream, Type type)
        {
            var workbook = WorkbookFactory.Create(stream);
            return Capture(workbook, type);
        }

        public object Capture(IWorkbook workbook, Type type)
        {
            var designClone = (TemplateDesign)_design.Clone();
            var obj = Read(workbook, type, designClone);

            if (_exceptions.Any())
            {
                throw _exceptions.First();
            }

            return obj;
        }

        public HintBuilder<T> CaptureHintBuilder<T>(Stream stream)
        {
            var workbook = WorkbookFactory.Create(stream);
            return CaptureHintBuilder<T>(workbook);
        }

        public HintBuilder<T> CaptureHintBuilder<T>(IWorkbook workbook)
        {
            var designClone = (TemplateDesign)_design.Clone();
            var obj = Read(workbook, typeof(T), designClone);

            if (_exceptions.Any())
            {
                throw _exceptions.First();
            }

            return new HintBuilder<T>(designClone, workbook, (T)obj, _exceptions);
        }

        /// <summary>
        /// 读取数据
        /// </summary>
        private object Read(IWorkbook workBook, Type type, TemplateDesign design)
        {
            _exceptions.Clear();

            var data = Activator.CreateInstance(type);
            var sheet = workBook.GetSheetAt(0);
            var current = design.BlockSection;
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

            foreach (var block in blocks.OfType<ValueBlock>())
            {
                var row = sheet.GetRow(block.Position.Row);
                var cell = row.GetCell(block.Position.Col, MissingCellPolicy.CREATE_NULL_AS_BLANK);

                try
                {
                    var rawVal = cell.GetValue();
                    var val = TryGetMappingValue(block.FieldPath, rawVal);
                    ObjectHelper.SetObjectValue(data, block.FieldPath, val);
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
                        throw new Exception($"只支持 List<T> 类型的集合，{table.TableName}");
                    }

                    int itemCount;
                    var list = ReadOneList(sheet, table.Body, prop.PropertyType, out itemCount);
                    var tableLastRow = table.Position.Row;
                    if (table.Header.Any())
                    {
                        tableLastRow = table.Header.Max(a => a.MergeTo?.Row ?? a.Position.Row);
                    }
                    else
                    {
                        tableLastRow = table.Body.Min(a => a.Position.Row);
                    }

                    tableLastRow += itemCount; //列表最后一行
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
                        var rawVal = cell.GetValue();
                        var val = TryGetMappingValue(block.FieldPath, rawVal);
                        var fieldPath = block.FieldPath.Substring(block.FieldPath.IndexOf('.') + 1);
                        ObjectHelper.SetObjectValue(obj, fieldPath, val);
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
        /// 添加值映射方法
        /// </summary>
        /// <param name="fieldPath"></param>
        /// <param name="mappingFunc"></param>
        /// <exception cref="Exception"></exception>
        public void AddMapping(string fieldPath, Func<object, object> mappingFunc)
        {
            if (_dicMappings.ContainsKey(fieldPath))
            {
                throw new Exception($"字段 {fieldPath} 已添加过 Mapping 方法了");
            }

            _dicMappings.Add(fieldPath, mappingFunc);
        }

        /// <summary>
        /// 尝试使用Mapping方法对值进行转化
        /// </summary>
        /// <param name="fieldPath"></param>
        /// <param name="rawVal"></param>
        /// <returns></returns>
        private object TryGetMappingValue(string fieldPath, object rawVal)
        {
            if (_dicMappings.TryGetValue(fieldPath, out Func<object, object> mappingFunc))
            {
                return mappingFunc(rawVal);
            }

            return rawVal;
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
