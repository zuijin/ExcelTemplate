using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ExcelTemplate.Attributes;
using ExcelTemplate.Helper;
using ExcelTemplate.Model;

namespace ExcelTemplate
{
    public static class TypeDesignAnalysis
    {
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

            var page = ReorganizePage(blocks);

            return new TemplateDesign(DesignSourceType.Object, page);
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

            if (string.IsNullOrWhiteSpace(valueAttr.Position) || !Position.IsPositionLetter(valueAttr.Position))
            {
                throw new Exception("集合的位置格式不正确");
            }

            var tablePosition = valueAttr.Position;
            var elementType = TypeHelper.GetCollectionElementType(prop.PropertyType);
            var subProps = elementType.GetProperties();
            var tmpList = new List<TypeRawHeader>();
            var dicMapper = new Dictionary<TableHeaderBlock, TableBodyBlock>();

            foreach (var subProp in subProps)
            {
                var path = PathCombine(parents?.Select(a => a.Name), prop.Name, subProp.Name);
                var propType = subProp.PropertyType;
                if (!TypeHelper.IsSimpleType(propType))
                {
                    throw new Exception($"集合内不支持任何复杂类型：{path}");
                }

                var titleAttr = subProp.GetCustomAttribute<TitleAttribute>();
                if (titleAttr != null)
                {
                    var headerBlock = new TableHeaderBlock()
                    {
                        Position = titleAttr.Position,
                        Text = titleAttr.Title,
                    };

                    //表体默认为表头的下一行
                    var bodyBlock = new TableBodyBlock()
                    {
                        Position = headerBlock.Position.GetOffset(1, 0),
                        FieldPath = path,
                    };

                    string[] mergeTitles = new string[0];
                    var mergeAttr = subProp.GetCustomAttribute<MergeAttribute>();
                    if (mergeAttr != null)
                    {
                        mergeTitles = mergeAttr.Titles;
                    }

                    dicMapper.Add(headerBlock, bodyBlock);
                    tmpList.Add(new TypeRawHeader() { Block = headerBlock, MergeTitles = mergeTitles });
                }
            }


            var headers = MergeHelper.MergeHeader(tablePosition, tmpList);
            foreach (var item in dicMapper)
            {
                item.Value.Position = item.Key.Position.GetOffset(1, 0);
            }

            var bodys = dicMapper.Select(item => item.Value).ToList();

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
        public static BlockPage ReorganizePage(List<IBlock> blocks)
        {
            if (blocks == null || !blocks.Any())
            {
                return null;
            }

            var groups = blocks.GroupBy(a => a.Position.Row).OrderBy(a => a.Key);
            var page = new BlockPage();
            var current = page;
            foreach (var rowBlocks in groups)
            {
                var tmp = new BlockPage()
                {
                    RowBlocks = rowBlocks.ToList(),
                };

                current.Next = tmp;
                current = tmp;
            }

            return page.Next;
        }

        private static string PathCombine(params string?[] paths)
        {
            return string.Join(".", paths.Where(a => !string.IsNullOrWhiteSpace(a)));
        }

        private static string PathCombine(IEnumerable<string?>? paths1, params string?[] paths2)
        {
            List<string?> paths = new List<string?>();
            if (paths1 != null)
            {
                paths.AddRange(paths1);
            }

            paths.AddRange(paths2);

            return string.Join(".", paths.Where(a => !string.IsNullOrWhiteSpace(a)));
        }
    }
}
