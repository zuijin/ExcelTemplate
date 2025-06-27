using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelTemplate.Attributes;
using ExcelTemplate.Helper;
using ExcelTemplate.Model;
using ExcelTemplate.Style;
using NPOI.OpenXmlFormats.Wordprocessing;

namespace ExcelTemplate
{
    public class TypeDesignAnalysis
    {
        /// <summary>
        /// 从类型中的字段特性（Attribute）定义，提取对应的模版设计信息
        /// </summary>
        /// <param name="type"></param>
        /// <param name="parents"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public TemplateDesign DesignAnalysis(Type type)
        {
            var blocks = new List<IBlock>();
            var styleDic = GetStyleDic(type);
            var props = type.GetProperties();

            foreach (PropertyInfo prop in props)
            {
                List<IBlock> tmpBlocks;
                var propType = prop.PropertyType;
                if (TypeHelper.IsSimpleType(propType))
                {
                    tmpBlocks = GetSimpleTypeBlocks(prop, styleDic);
                }
                else if (IsSpecialCollectionType(propType))
                {
                    tmpBlocks = GetCollectionTypeBlocks(prop, styleDic);
                }
                else
                {
                    throw new Exception($"暂不支持复杂类型：{prop.Name}");
                }

                blocks.AddRange(tmpBlocks);
            }

            var section = ReorganizeSection(blocks);

            //合并
            MergeSection(section);

            return new TemplateDesign(TemplateDesignSourceType.Type, section);
        }

        private static Dictionary<string, IETStyle> GetStyleDic(Type type)
        {
            var styleDic = new Dictionary<string, IETStyle>();
            var attrs = type.GetCustomAttributes<StyleDicAttribute>();

            foreach (var attr in attrs)
            {
                var style = ETStyleUtil.ConvertStyle(attr);
                styleDic.Add(attr.Key, style);
            }

            return styleDic;
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
        private static List<IBlock> GetSimpleTypeBlocks(PropertyInfo prop, Dictionary<string, IETStyle> dicStyle)
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
                    Style = GetBlockStyle(titleAttr.Style, dicStyle),
                });
            }

            var valueAttr = prop.GetCustomAttribute<PositionAttribute>();
            if (valueAttr != null)
            {
                blocks.Add(new ValueBlock()
                {
                    Position = valueAttr.Position,
                    MergeTo = valueAttr.MergeTo,
                    FieldPath = prop.Name,
                    Style = GetBlockStyle(valueAttr.Style, dicStyle, prop),
                });
            }

            return blocks;
        }

        /// <summary>
        /// 获取区块样式
        /// </summary>
        /// <param name="styleKey"></param>
        /// <param name="dicStyle"></param>
        /// <param name="prop"></param>
        /// <returns></returns>
        private static IETStyle GetBlockStyle(string styleKey, Dictionary<string, IETStyle> dicStyle, PropertyInfo prop = null)
        {
            IETStyle style = null;
            if (!string.IsNullOrWhiteSpace(styleKey))
            {
                dicStyle.TryGetValue(styleKey, out style);
            }
            else if (prop != null)
            {
                var attr = prop.GetCustomAttribute<StyleAttribute>();
                if (attr != null)
                {
                    style = ETStyleUtil.ConvertStyle(attr);
                }
            }

            return style;
        }

        /// <summary>
        /// 获取集合类型Block
        /// </summary>
        /// <param name="prop"></param>
        /// <param name="parent"></param>
        /// <returns></returns>
        private static List<IBlock> GetCollectionTypeBlocks(PropertyInfo prop, Dictionary<string, IETStyle> dicStyle)
        {
            List<IBlock> blocks = new List<IBlock>();
            var positionAttr = prop.GetCustomAttribute<PositionAttribute>();
            if (positionAttr == null)
            {
                return blocks;
            }

            var titleAttrs = prop.GetCustomAttributes<TitleAttribute>();
            foreach (var attr in titleAttrs)
            {
                blocks.Add(new TextBlock()
                {
                    Text = attr.Title,
                    Position = attr.Position,
                    MergeTo = attr.MergeTo,
                    Style = GetBlockStyle(attr.Style, dicStyle),
                });
            }

            var tablePosition = positionAttr.Position;
            var elementType = TypeHelper.GetCollectionElementType(prop.PropertyType);
            var subProps = elementType.GetProperties();
            var rawHeaderList = new List<TypeRawHeader>();
            var bodys = new List<TableBodyBlock>();
            var headStyle = GetBlockStyle(positionAttr.Style, dicStyle, prop);

            foreach (var subProp in subProps)
            {
                var path = PathCombine(prop.Name, subProp.Name);
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
                        Position = positionAttr.Position.GetOffset(0, colAttr.ColIndex),
                        Text = colAttr.HeaderText,
                        Style = headStyle,
                    };

                    //表体默认为表头的下一行
                    var bodyBlock = new TableBodyBlock()
                    {
                        Position = positionAttr.Position.GetOffset(1, colAttr.ColIndex),
                        FieldPath = path,
                        Style = GetBlockStyle(colAttr.Style, dicStyle, subProp),
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

            var headers = MergeHelper.MergeHeader(tablePosition, rawHeaderList, headStyle);
            foreach (var body in bodys)
            {
                // 对应那一列最低的一个表头
                var lowestHeader = headers.Where(a => a.Position.Col == body.Position.Col)
                        .OrderByDescending(a => a.Position.Row)
                        .First();

                body.Position.Row = (lowestHeader.MergeTo ?? lowestHeader.Position).Row + 1;
            }

            blocks.Add(new TableBlock()
            {
                TableName = prop.Name,
                Position = tablePosition,
                Header = headers,
                Body = bodys,
            });

            return blocks;
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
                else
                {
                    preSection = current;
                }

                current = current.Next;
                preIsTable = isTable;
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
