using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using ExcelTemplate.Exceptions;
using ExcelTemplate.Extensions;
using ExcelTemplate.Helper;
using ExcelTemplate.Hint;
using ExcelTemplate.Model;
using ExcelTemplate.Style;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelTemplate
{
    /// <summary>
    /// 负责按照模版定义，把数据写入到Excel中
    /// </summary>
    public class TemplateRender
    {
        TemplateDesign _design;

        public TemplateDesign Design { get => _design; }
        Dictionary<string, Func<object, object>> _dicMappings;

        /// <summary>
        /// 
        /// </summary>
        public TemplateRender(TemplateDesign design)
        {
            _design = design;
            _dicMappings = new Dictionary<string, Func<object, object>>();

            DesignInspector.Check(design);
        }

        /// <summary>
        /// 创建 TemplateWriter
        /// </summary>
        public static TemplateRender Create(Type type)
        {
            var design = TypeDesignAnalysis.DesignAnalysis(type);
            return new TemplateRender(design);
        }

        private static IWorkbook CreateWorkbook(ExcelType excelType)
        {
            IWorkbook workbook;
            if (excelType == ExcelType.Xlsx)
            {
                workbook = new XSSFWorkbook();
            }
            else if (excelType == ExcelType.Xls)
            {
                workbook = new HSSFWorkbook();
            }
            else
            {
                throw new Exception("不支持的 ExcelType 枚举");
            }

            return workbook;
        }

        /// <summary>
        /// 渲染数据
        /// </summary>
        /// <returns></returns>
        public IWorkbook Render(object data, ExcelType excelType = ExcelType.Xlsx)
        {
            var workbook = CreateWorkbook(excelType);
            Render(data, workbook);

            return workbook;
        }

        public void Render(object data, IWorkbook workbook)
        {
            if (workbook.NumberOfSheets == 0)
            {
                workbook.CreateSheet();
            }

            var designClone = (TemplateDesign)_design.Clone();
            Write(designClone, workbook, data);
        }

        public HintBuilder<T> RenderHintBuilder<T>(T data, ExcelType excelType = ExcelType.Xlsx)
        {
            var workbook = CreateWorkbook(excelType);
            return RenderHintBuilder<T>(data, workbook);
        }

        public HintBuilder<T> RenderHintBuilder<T>(T data, IWorkbook workbook)
        {
            if (workbook.NumberOfSheets == 0)
            {
                workbook.CreateSheet();
            }

            var designClone = (TemplateDesign)_design.Clone();
            Write(designClone, workbook, data);

            return new HintBuilder<T>(designClone, workbook, data, new List<CellException> { });
        }

        /// <summary>
        /// 写入表单数据
        /// </summary>
        protected void Write(TemplateDesign design, IWorkbook workbook, object data)
        {
            var sheet = workbook.GetSheetAt(0);
            var current = design.BlockSection;

            while (current != null)
            {
                if (DesignInspector.IsTableSection(current))
                {
                    WriteTable(current, sheet, data);
                }
                else
                {
                    WriteForm(current, sheet, data);
                }

                current = current.Next;
            }
        }

        /// <summary>
        /// 写入表单数据
        /// </summary>
        /// <param name="section"></param>
        /// <param name="sheet"></param>
        /// <param name="data"></param>
        protected void WriteForm(BlockSection section, ISheet sheet, object data)
        {
            foreach (var textBlock in section.Blocks.OfType<TextBlock>())
            {
                var cell = sheet.GetOrCreateRow(textBlock.Position.Row).GetOrCreateCell(textBlock.Position.Col);
                cell.SetValue(textBlock.Text);
                cell.SetStyle(textBlock.Style);

                if (textBlock.MergeTo != null)
                {
                    SetMergeStyle(sheet, textBlock.Position, textBlock.MergeTo, textBlock.Style);
                    sheet.AddMergedRegion(textBlock.Position, textBlock.MergeTo);
                }
            }

            foreach (var valueBlock in section.Blocks.OfType<ValueBlock>())
            {
                var cell = sheet.GetOrCreateRow(valueBlock.Position.Row).GetOrCreateCell(valueBlock.Position.Col);
                var val = ObjectHelper.GetObjectValue(data, valueBlock.FieldPath);
                var cellVal = TryGetMappingValue(valueBlock.FieldPath, val);
                cell.SetValue(cellVal);
                cell.SetStyle(valueBlock.Style);

                if (valueBlock.MergeTo != null)
                {
                    SetMergeStyle(sheet, valueBlock.Position, valueBlock.MergeTo, valueBlock.Style);
                    sheet.AddMergedRegion(valueBlock.Position, valueBlock.MergeTo);
                }
            }
        }

        /// <summary>
        /// 读取列表数据
        /// </summary>
        /// <param name="section"></param>
        /// <param name="sheet"></param>
        /// <param name="data"></param>
        /// <returns>返回最后一行的下标</returns>
        /// <exception cref="Exception"></exception>
        protected void WriteTable(BlockSection section, ISheet sheet, object data)
        {
            var tables = section.Blocks.OfType<TableBlock>();
            var props = data.GetType().GetProperties();
            var nextOffset = 0;

            foreach (var table in tables)
            {
                var prop = props.FirstOrDefault(a => a.Name == table.TableName);
                if (prop != null)
                {
                    if (!TypeHelper.IsSubclassOfRawGeneric(typeof(List<>), prop.PropertyType))
                    {
                        throw new Exception("只支持 List<T> 类型的集合");
                    }

                    var list = prop.GetValue(data);
                    int count = WriteOneList(sheet, table, list);
                    nextOffset = Math.Max(nextOffset, count - 1);
                }
            }

            //后面的区块自动偏移
            section.Next?.ApplyOffset(nextOffset, 0);
        }

        /// <summary>
        /// 写入列表数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="table"></param>
        /// <param name="listData"></param>
        /// <returns>返回数据行数</returns>
        protected int WriteOneList(ISheet sheet, TableBlock table, object listData)
        {
            // 写入表头
            foreach (var header in table.Header)
            {
                var cell = sheet.GetOrCreateRow(header.Position.Row).GetOrCreateCell(header.Position.Col);
                cell.SetValue(header.Text);
                cell.SetStyle(header.Style);

                if (header.MergeTo != null)
                {
                    SetMergeStyle(sheet, header.Position, header.MergeTo, header.Style);
                    sheet.AddMergedRegion(header.Position, header.MergeTo);
                }
            }

            // 写入表体数据
            if (listData != null && listData is IEnumerable list)
            {
                int firstRow = table.Body.First().Position.Row;
                int rowIndex = firstRow;
                foreach (var item in list)
                {
                    var itemProps = item.GetType().GetProperties();
                    var row = sheet.GetOrCreateRow(rowIndex);
                    foreach (var body in table.Body)
                    {
                        var filePath = body.FieldPath.Substring(body.FieldPath.LastIndexOf('.') + 1);
                        var val = ObjectHelper.GetObjectValue(item, filePath);
                        var cell = row.GetOrCreateCell(body.Position.Col);
                        var cellVal = TryGetMappingValue(body.FieldPath, val);
                        cell.SetValue(cellVal);
                        cell.SetStyle(body.Style);
                    }

                    rowIndex++;
                }

                return (rowIndex - firstRow);
            }

            return 0;
        }

        protected void SetMergeStyle(ISheet sheet, Position position, Position mergeTo, IETStyle style)
        {
            if (style != null)
            {
                for (int i = 0; i <= (mergeTo.Row - position.Row); i++)
                {
                    for (int j = 0; j <= (mergeTo.Col - position.Col); j++)
                    {
                        var pos = position.GetOffset(i, j);
                        sheet.GetOrCreateCell(pos).SetStyle(style);
                    }
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
    }
}
