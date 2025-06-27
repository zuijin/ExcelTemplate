using ExcelTemplate.Extensions;
using ExcelTemplate.Helper;
using ExcelTemplate.Model;
using ExcelTemplate.Style;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelTemplate
{
    public class ExcelDesignAnalysis
    {
        private const string VALUE_FIELD_FORMAT = @"^\${(([_a-zA-Z][_a-zA-Z0-9]*)(\.[_a-zA-Z][_a-zA-Z0-9]*)*)}$";
        private const string VALUE_FIELD_PATH = @"^\${(.*)}$";

        private const string TBODY_FIELD_FORMAT = @"^\${#(([_a-zA-Z][_a-zA-Z0-9]*)(\.[_a-zA-Z][_a-zA-Z0-9]*)*)}$";
        private const string TBODY_FIELD_PATH = @"^\${#(.*)}$";

        private List<IETStyle> _uniqueStyles = new List<IETStyle>();

        public IETStyle GetOrMapStyle(ICell cell)
        {
            var style = ETStyleUtil.ConvertStyle(cell.Sheet.Workbook, cell.CellStyle);
            foreach (var item in _uniqueStyles)
            {
                if (ObjectHelper.Compare(style, item))
                {
                    return item;
                }
            }

            _uniqueStyles.Add(style);
            return style;
        }

        /// <summary>
        /// 从excel文件中提取对应的模版设计信息
        /// </summary>
        /// <param name="firstSection"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception> 
        public TemplateDesign DesignAnalysis(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName) || !File.Exists(fileName))
            {
                throw new Exception($"文件{fileName}不存在");
            }

            var workbook = WorkbookFactory.Create(fileName);
            var sheet = workbook.GetSheetAt(0);
            var rowEnumerator = sheet.GetRowEnumerator();
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

                        if (valueStr != null && Regex.IsMatch(valueStr, VALUE_FIELD_FORMAT))
                        {
                            blocks.Add(new ValueBlock()
                            {
                                FieldPath = Regex.Match(valueStr, VALUE_FIELD_PATH).Groups[1].Value,
                                Position = position,
                                Style = style,
                                MergeTo = mergeTo,
                            });
                        }
                        else if (valueStr != null && Regex.IsMatch(valueStr, TBODY_FIELD_FORMAT))
                        {
                            blocks.Add(new TableBodyBlock()
                            {
                                FieldPath = Regex.Match(valueStr, TBODY_FIELD_PATH).Groups[1].Value,
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

            return new TemplateDesign(TemplateDesignSourceType.File, firstSection);
        }

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

        private string GetTableName(string fieldPath)
        {
            return fieldPath.Substring(0, fieldPath.LastIndexOf('.'));
        }
    }
}
