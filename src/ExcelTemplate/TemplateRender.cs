using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ExcelTemplate.Extensions;
using ExcelTemplate.Helper;
using ExcelTemplate.Model;
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

        /// <summary>
        /// 
        /// </summary>
        public TemplateRender(TemplateDesign design)
        {
            _design = design;
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

        /// <summary>
        /// 写入数据
        /// </summary>
        /// <returns></returns>
        public IWorkbook Render(object data)
        {
            var workbook = new XSSFWorkbook();
            workbook.CreateSheet();

            Write(_design, workbook, data);

            return workbook;
        }

        public void Render(object data, IWorkbook workbook)
        {
            if (workbook.NumberOfSheets == 0)
            {
                workbook.CreateSheet();
            }

            Write(_design, workbook, data);
        }

        /// <summary>
        /// 读取表单数据
        /// </summary>
        protected void Write(TemplateDesign design, IWorkbook workbook, object data)
        {
            var sheet = workbook.GetSheetAt(0);
            var current = (BlockSection)design.BlockSection.Clone();

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

                cell.SetCellValue(textBlock.Text);

                if (textBlock.MergeTo != null)
                {
                    sheet.AddMergedRegion(textBlock.Position, textBlock.MergeTo);
                }
            }

            foreach (var valueBlock in section.Blocks.OfType<ValueBlock>())
            {
                var cell = sheet.GetOrCreateRow(valueBlock.Position.Row).GetOrCreateCell(valueBlock.Position.Col);
                var val = ObjectHelper.GetObjectValue(data, valueBlock.FieldPath);

                cell.SetCellValue(val.ToString());

                if (valueBlock.MergeTo != null)
                {
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
                var tableHeight = table.Header.Max(a => a.MergeTo?.Row ?? a.Position.Row) - table.Position.Row;
                var prop = props.FirstOrDefault(a => a.Name == table.TableName);
                if (prop != null)
                {
                    if (!TypeHelper.IsSubclassOfRawGeneric(typeof(List<>), prop.PropertyType))
                    {
                        throw new Exception("只支持 List<T> 类型的集合");
                    }

                    var list = prop.GetValue(data);
                    int count = WriteOneList(sheet, table, list);
                    tableHeight += count;
                }

                nextOffset = Math.Max(nextOffset, tableHeight);
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
                cell.SetCellValue(header.Text);

                if (header.MergeTo != null)
                {
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
                        cell.SetCellValue(val?.ToString());
                    }

                    rowIndex++;
                }

                return (rowIndex - firstRow);
            }

            return 0;
        }
    }
}
