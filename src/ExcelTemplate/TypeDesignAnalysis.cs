using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using ExcelTemplate.Attributes;
using ExcelTemplate.Helper;
using ExcelTemplate.Model;

namespace ExcelTemplate
{
    public static class TypeDesignAnalysis
    {
        public static TemplateDesign DesignAnalysis(Type type, List<PropertyInfo>? parents = null)
        {
            var blocks = new TemplateDesign();
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
                    blocks.AddRange(GetCollectionTypeBlocks(prop, parents));
                }
                else
                {
                    throw new Exception($"暂不支持复杂类型：{prop.Name}");
                }
            }

            return blocks;
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
        private static List<Block> GetSimpleTypeBlocks(PropertyInfo prop, List<PropertyInfo>? parents = null)
        {
            var blocks = new List<Block>();
            var titleAttr = prop.GetCustomAttribute<TitleAttribute>();
            if (titleAttr != null)
            {
                blocks.Add(new Block()
                {
                    BlockType = BlockType.Constant,
                    Position = titleAttr.Position,
                    Value = titleAttr.Title,
                    MergeTo = titleAttr.MergeTo,
                });
            }

            var valueAttr = prop.GetCustomAttribute<ValueAttribute>();
            if (valueAttr != null)
            {
                blocks.Add(new Block()
                {
                    BlockType = BlockType.Variable,
                    Position = valueAttr.Position,
                    MergeTo = valueAttr.MergeTo,
                    ValuePath = PathCombine(parents?.Select(a => a.Name), prop.Name)
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
        private static List<Block> GetCollectionTypeBlocks(PropertyInfo prop, List<PropertyInfo>? parents = null)
        {
            var blocks = new List<Block>();
            var valueAttr = prop.GetCustomAttribute<ValueAttribute>();
            if (valueAttr == null)
            {
                return blocks;
            }

            if (string.IsNullOrWhiteSpace(valueAttr.Position) || !Position.IsPositionLetter(valueAttr.Position))
            {
                throw new Exception("集合的位置格式不正确");
            }

            var elementType = TypeHelper.GetCollectionElementType(prop.PropertyType);
            var subProps = elementType.GetProperties();
            var tmpList = new List<HeaderBlock>();

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
                    var block = new Block()
                    {
                        BlockType = BlockType.Array,
                        Position = titleAttr.Position,
                        ValuePath = path,
                        Value = titleAttr.Title,
                    };

                    string[] mergeTitles = new string[0];
                    var mergeAttr = subProp.GetCustomAttribute<MergeAttribute>();
                    if (mergeAttr != null)
                    {
                        mergeTitles = mergeAttr.Titles;
                    }

                    tmpList.Add(new HeaderBlock() { Block = block, MergeTitles = mergeTitles });
                }
            }

            blocks = MergeHelper.MergeHeader(valueAttr.Position, tmpList);

            return blocks;
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
