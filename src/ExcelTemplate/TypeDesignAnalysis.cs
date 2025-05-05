using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelTemplate.Attributes;
using ExcelTemplate.Helper;
using ExcelTemplate.Model;
using NPOI.SS.Formula.Functions;

namespace ExcelTemplate
{
    public static class TypeDesignAnalysis
    {
        /// <summary>
        /// 从类型中的字段特性（Attribute）定义，提取对应的模版设计信息
        /// </summary>
        /// <param name="type"></param>
        /// <param name="parents"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static TemplateDesign DesignAnalysis(Type type, List<PropertyInfo>? parents = null)
        {
            var blocks = new List<IBlock>();
            var props = type.GetProperties();
            foreach (PropertyInfo prop in props)
            {
                var propType = prop.PropertyType;
                if (TypeHelper.IsSimpleType(propType))
                {
                    blocks.AddRange(GetSimpleTypeBlocks(prop, parents));
                }
                else if (IsSpecialCollectionType(propType))
                {
                    blocks.Add(GetCollectionTypeBlocks(prop, parents));
                }
                else
                {
                    throw new Exception($"暂不支持复杂类型：{prop.Name}");
                }
            }

            var section = ReorganizeSection(blocks);

            //合并
            MergeSection(section);

            return new TemplateDesign(DesignSourceType.Object, section);
        }


        /// <summary>
        /// 判断是否支持的集合类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        private static bool IsSpecialCollectionType(Type type)
        {
            if (!TypeHelper.IsCollectionType(type))
                return false;

            // 获取集合元素的类型
            Type elementType = TypeHelper.GetCollectionElementType(type);

            // 判断元素类型是否为简单类型
            return !TypeHelper.IsSimpleType(elementType) && !TypeHelper.IsCollectionType(elementType);
        }

        /// <summary>
        /// 获取简单类型Block
        /// </summary>
        /// <param name="prop"></param>
        /// <param name="parent"></param>
        /// <returns></returns>
        private static List<IBlock> GetSimpleTypeBlocks(PropertyInfo prop, List<PropertyInfo>? parents = null)
        {
            var blocks = new List<IBlock>();
            var titleAttr = prop.GetCustomAttribute<TitleAttribute>();
            if (titleAttr != null)
            {
                blocks.Add(new TextBlock()
                {
                    Position = titleAttr.Position,
                    Text = titleAttr.Title,
                    MergeTo = titleAttr.MergeTo,
                });
            }

            var valueAttr = prop.GetCustomAttribute<PositionAttribute>();
            if (valueAttr != null)
            {
                blocks.Add(new ValueBlock()
                {
                    Position = valueAttr.Position,
                    MergeTo = valueAttr.MergeTo,
                    FieldPath = PathCombine(parents?.Select(a => a.Name), prop.Name)
                });
            }

            return blocks;
        }

        /// <summary>
        /// 获取集合类型Block
        /// </summary>
        /// <param name="prop"></param>
        /// <param name="parent"></param>
        /// <returns></returns>
        private static TableBlock GetCollectionTypeBlocks(PropertyInfo prop, List<PropertyInfo>? parents = null)
        {
            var valueAttr = prop.GetCustomAttribute<PositionAttribute>();
            if (valueAttr == null)
            {
                return null;
            }

            var tablePosition = valueAttr.Position;
            var elementType = TypeHelper.GetCollectionElementType(prop.PropertyType);
            var subProps = elementType.GetProperties();
            var rawHeaderList = new List<TypeRawHeader>();
            var bodys = new List<TableBodyBlock>();

            foreach (var subProp in subProps)
            {
                var path = PathCombine(parents?.Select(a => a.Name), prop.Name, subProp.Name);
                var propType = subProp.PropertyType;
                if (!TypeHelper.IsSimpleType(propType))
                {
                    throw new Exception($"集合内不支持任何复杂类型：{path}");
                }

                var colAttr = subProp.GetCustomAttribute<ColAttribute>();
                if (colAttr != null)
                {
                    var headerBlock = new TableHeaderBlock()
                    {
                        Position = colAttr.Position,
                        Text = colAttr.HeaderText,
                    };

                    //表体默认为表头的下一行
                    var bodyBlock = new TableBodyBlock()
                    {
                        Position = headerBlock.Position.GetOffset(1, 0),
                        FieldPath = path,
                    };

                    string[] mergeTitles = new string[0];
                    var mergeAttr = subProp.GetCustomAttribute<MergeAttribute>();
                    if (mergeAttr != null && mergeAttr.Titles != null)
                    {
                        mergeTitles = mergeAttr.Titles;
                    }

                    rawHeaderList.Add(new TypeRawHeader() { Block = headerBlock, MergeTitles = mergeTitles });
                    bodys.Add(bodyBlock);
                }
            }

            var headers = MergeHelper.MergeHeader(tablePosition, rawHeaderList);
            foreach (var body in bodys)
            {
                // 对应那一列最低的一个表头
                var lowestHeader = headers.Where(a => a.Position.Col == body.Position.Col)
                        .OrderByDescending(a => a.Position.Row)
                        .First();

                body.Position.Row = (lowestHeader.MergeTo ?? lowestHeader.Position).Row + 1;
            }

            return new TableBlock()
            {
                TableName = PathCombine(parents?.Select(a => a.Name), prop.Name),
                Position = tablePosition,
                Header = headers,
                Body = bodys,
            };
        }

        /// <summary>
        /// 重新整理Block顺序，按照层叠式结构处理
        /// </summary>
        /// <param name="blocks"></param>
        /// <returns></returns>
        public static BlockSection ReorganizeSection(List<IBlock> blocks)
        {
            if (blocks == null || !blocks.Any())
            {
                return null;
            }

            var groups = blocks.GroupBy(a => a.Position.Row).OrderBy(a => a.Key);
            var section = new BlockSection();
            var current = section;
            foreach (var rowBlocks in groups)
            {
                var tmp = new BlockSection()
                {
                    Blocks = rowBlocks.ToList(),
                };

                current.Next = tmp;
                current = tmp;
            }

            return section.Next;
        }

        /// <summary>
        /// 将部分可以合并的区块，尽量合并起来
        /// </summary>
        /// <param name="section"></param>
        /// <returns></returns>
        public static void MergeSection(BlockSection section)
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

                current = current.Next;
            }
        }

        private static string PathCombine(params string[] paths)
        {
            return string.Join(".", paths.Where(a => !string.IsNullOrWhiteSpace(a)));
        }

        private static string PathCombine(IEnumerable<string> paths1, params string[] paths2)
        {
            List<string> paths = new List<string>();
            if (paths1 != null)
            {
                paths.AddRange(paths1);
            }

            paths.AddRange(paths2);

            return string.Join(".", paths.Where(a => !string.IsNullOrWhiteSpace(a)));
        }
    }
}
