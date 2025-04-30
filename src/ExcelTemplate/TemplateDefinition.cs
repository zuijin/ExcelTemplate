using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTemplate
{
    internal class TemplateDefinition : List<Block>
    {
        public ExcelType ExcelType { get; set; }
    }

    public enum ExcelType
    {
        Xls = 0,
        Xlsx = 1,
    }

    public class Block
    {
        /// <summary>
        /// 类型
        /// </summary>
        public BlockType BlockType { get; set; }

        public object? Value { get; set; }

        public string? ValuePath { get; set; }

        public Position? Position { get; set; }

        public Position? MergeTo { get; set; }
    }

    public enum BlockType
    {
        /// <summary>
        /// 常量
        /// </summary>
        Constant = 0,
        /// <summary>
        /// 变量
        /// </summary>
        Variable = 1,
        /// <summary>
        /// 对象
        /// </summary>
        Object = 3,
        /// <summary>
        /// 数组
        /// </summary>
        Array = 4
    }
}
