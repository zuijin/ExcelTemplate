using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelTemplate.Extensions;
using ExcelTemplate.Model;
using ExcelTemplate.Style;
using NPOI.SS.UserModel;

namespace ExcelTemplate
{
    public class ExcelDesignAnalysis
    {
        private static readonly Regex ValueFieldRegex = new Regex(@"^\${(([_a-zA-Z][_a-zA-Z0-9]*)(\.[_a-zA-Z][_a-zA-Z0-9]*)*)}$", RegexOptions.Compiled);
        private static readonly Regex TBodyFieldRegex = new Regex(@"^\${#(([_a-zA-Z][_a-zA-Z0-9]*)(\.[_a-zA-Z][_a-zA-Z0-9]*)*)}$", RegexOptions.Compiled);

        private readonly Dictionary<IETStyle, IETStyle> _uniqueStyles = new Dictionary<IETStyle, IETStyle>();

        /// <summary>
        /// 获取或映射单元格样式
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <returns>样式对象</returns>
        public IETStyle GetOrMapStyle(ICell cell)
        {
            var style = ETStyleUtil.ConvertStyle(cell.Sheet.Workbook, cell.CellStyle);
            if (_uniqueStyles.TryGetValue(style, out var existingStyle))
            {
                return existingStyle;
            }

            _uniqueStyles[style] = style;
            return style;
        }

        /// <summary>
        /// 从 Excel 文件中提取对应的模版设计信息
        /// </summary>
        /// <param name="fileName">文件路径</param>
        /// <returns>模版设计信息</returns>
        public TemplateDesign DesignAnalysis(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName) || !File.Exists(fileName))
            {
                throw new Exception($"文件{fileName}不存在");
            }

            using (var s = File.OpenRead(fileName))
            {
                return DesignAnalysis(s);
            }
        }

        /// <summary>
        /// 从excel文件中提取对应的模版设计信息
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public TemplateDesign DesignAnalysis(Stream stream)
        {
            var workbook = WorkbookFactory.Create(stream);
            var sheet = workbook.GetSheetAt(0);
            var rowEnumerator = sheet.GetEnumerator();
            var mergeInfos = sheet.MergedRegions;
            BlockSection firstSection = null;
            BlockSection currentSection = null;

            while (rowEnumerator.MoveNext())
            {
                var cellEnumerator = ((IRow)rowEnumerator.Current).GetEnumerator();
                var blocks = new List<IBlock>();

                while (cellEnumerator.MoveNext())
                {
                    var cell = cellEnumerator.Current;
                    if (!cell.IsEmpty())
                    {
                        var val = cell.GetValue();
                        var position = new Position(cell.RowIndex, cell.ColumnIndex);
                        var merge = mergeInfos.Find(a => a.FirstColumn == cell.ColumnIndex && a.FirstRow == cell.RowIndex);
                        var mergeTo = (merge == null) ? null : new Position(merge.LastRow, merge.LastColumn);
                        var style = GetOrMapStyle(cell);
                        var valueStr = val is string ? val.ToString().Replace(" ", "") : null;
                        Match valueMatch = null;
                        Match tbodyMatch = null;

                        if (valueStr != null && (valueMatch = ValueFieldRegex.Match(valueStr)).Success)
                        {
                            blocks.Add(new ValueBlock()
                            {
                                FieldPath = valueMatch.Groups[1].Value,
                                Position = position,
                                Style = style,
                                MergeTo = mergeTo,
                            });
                        }
                        else if (valueStr != null && (tbodyMatch = TBodyFieldRegex.Match(valueStr)).Success)
                        {
                            blocks.Add(new TableBodyBlock()
                            {
                                FieldPath = tbodyMatch.Groups[1].Value,
                                Position = position,
                                Style = style,
                                MergeTo = mergeTo,
                            });
                        }
                        else
                        {
                            blocks.Add(new TextBlock()
                            {
                                Text = val.ToString(),
                                Position = position,
                                Style = style,
                                MergeTo = mergeTo,
                            });
                        }
                    }
                }

                if (blocks.Any())
                {
                    var section = new BlockSection() { Blocks = blocks };
                    if (currentSection == null)
                    {
                        currentSection = section;
                        firstSection = section;
                    }
                    else
                    {
                        currentSection.Next = section;
                        currentSection = section;
                    }
                }
            }

            if (firstSection == null)
            {
                return null;
            }

            CheckAndBuildTable(firstSection);

            MergeSection(firstSection);

            return new TemplateDesign(TemplateDesignSourceType.File, firstSection);
        }

        /// <summary>
        /// 检查并构建表格区块
        /// </summary>
        /// <param name="root">根区块段</param>
        private void CheckAndBuildTable(BlockSection root)
        {
            var current = root;
            while (current != null)
            {
                if (current.Blocks.Any(a => a is TableBodyBlock))
                {
                    var tableBodys = current.Blocks.OfType<TableBodyBlock>().OrderBy(a => a.Position.Col).ToList();
                    current.Blocks.RemoveAll(a => a is TableBodyBlock);

                    var tableGroups = tableBodys.GroupBy(a => GetTableName(a.FieldPath));
                    foreach (var tableGroup in tableGroups)
                    {
                        current.Blocks.Add(new TableBlock()
                        {
                            Body = tableBodys,
                            Header = new List<TableHeaderBlock>(),
                            TableName = tableGroup.Key,
                            Position = tableBodys.First().Position,
                        });
                    }
                }

                current = current.Next;
            }
        }

        /// <summary>
        /// 获取表名
        /// </summary>
        /// <param name="fieldPath">字段路径</param>
        /// <returns>表名</returns>
        private string GetTableName(string fieldPath)
        {
            return fieldPath.Substring(0, fieldPath.LastIndexOf('.'));
        }

        /// <summary>
        /// 将部分可以合并的区块，尽量合并起来
        /// </summary>
        /// <param name="section"></param>
        /// <returns></returns>
        private static void MergeSection(BlockSection section)
        {
            var preSection = section;
            var preIsTable = DesignInspector.IsTableSection(section);
            var current = section.Next;

            while (current != null)
            {
                var isTable = DesignInspector.IsTableSection(current);
                if (!isTable && !preIsTable) // 两个非列表区块，可以合并
                {
                    preSection.Blocks.AddRange(current.Blocks);
                    preSection.Next = current.Next;
                }
                else
                {
                    preSection = current;
                }

                current = current.Next;
                preIsTable = isTable;
            }
        }
    }
}
