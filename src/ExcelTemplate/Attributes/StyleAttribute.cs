using System;
using System.Collections.Generic;
using System.Text;
using ExcelTemplate.Style;

namespace ExcelTemplate.Attributes
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class StyleAttribute : Attribute, ISimpleStyle
    {
        /// <summary>
        /// 获取或设置数据格式索引
        /// </summary>
        public string DataFormat { get; set; }

        /// <summary>
        /// 前景色
        /// </summary>
        public string ForegroundColor { get; set; }

        /// <summary>
        /// 背景色
        /// </summary>
        public string BackgroundColor { get; set; }

        /// <summary>
        /// 获取或设置是否使用粗体
        /// </summary>
        public bool IsBold { get; set; }

        /// <summary>
        /// 获取或设置字体高度（以磅为单位的数值）
        /// </summary>
        /// <remarks>
        /// 此属性返回的值与Excel中显示的字号一致，如10、14或28等
        /// </remarks>
        /// <see cref="FontHeight"/>
        public double FontHeightInPoints { get; set; }

        /// <summary>
        /// 单元格是否自动伸缩以适应文本（当文本过长时）
        /// </summary>
        public bool ShrinkToFit { get; set; }

        /// <summary>
        /// 获取或设置水平对齐方式
        /// </summary>
        public HorizontalAlignment Alignment { get; set; }

        /// <summary>
        /// 获取或设置是否自动换行
        /// </summary>
        public bool WrapText { get; set; }

        /// <summary>
        /// 获取或设置垂直对齐方式
        /// </summary>
        public VerticalAlignment VerticalAlignment { get; set; }
    }

    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
    public class StyleDicAttribute : StyleAttribute
    {
        /// <summary>
        /// 为样式设置一个唯一的 Key
        /// </summary>
        public string Key { get; set; }
    }
}
