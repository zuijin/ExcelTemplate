using System.ComponentModel.DataAnnotations;
using System.Reflection;
using ExcelTemplate.Attributes;

namespace ExcelTemplate.Model
{
    /// <summary>
    /// excel字段定义
    /// </summary>
    public class FieldInfo
    {
        /// <summary>
        /// 显示名称
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// 标题位置
        /// </summary>
        public Position TitlePosition { get; set; }
        /// <summary>
        /// 标题合并结束位置
        /// </summary>
        public Position? TitleMergeTo { get; set; }
        /// <summary>
        /// 值位置
        /// </summary>
        public Position ValuePosition { get; set; }
        /// <summary>
        /// 值合并结束位置
        /// </summary>
        public Position ValueMergeTo { get; set; }
        /// <summary>
        /// 表头合并
        /// </summary>
        public string[] HeaderMerge { get; set; }

        /// <summary>
        /// 对应的字段属性
        /// </summary>
        public PropertyInfo Property { get; set; }
    }
}
