using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTemplate.Style
{
    /// <summary>
    /// 简单样式
    /// </summary>
    public interface IETSimpleStyle
    {
        /// <summary>
        /// 获取或设置数据格式索引
        /// </summary>
        string DataFormat { get; set; }

        /// <summary>
        /// 前景色
        /// </summary>
        string TextColor { get; set; }

        /// <summary>
        /// 背景色
        /// </summary>
        string BgColor { get; set; }

        /// <summary>
        /// 获取或设置是否使用粗体
        /// </summary>
        bool IsBold { get; set; }

        /// <summary>
        /// 获取或设置字体高度（以磅为单位的数值）
        /// </summary>
        /// <remarks>
        /// 此属性返回的值与Excel中显示的字号一致，如10、14或28等
        /// </remarks>
        /// <see cref="FontHeight"/>
        double FontHeightInPoints { get; set; }

        /// <summary>
        /// 单元格是否自动伸缩以适应文本（当文本过长时）
        /// </summary>
        bool ShrinkToFit { get; set; }

        /// <summary>
        /// 获取或设置水平对齐方式
        /// </summary>
        ETHorizontalAlignment HorizontalAlignment { get; set; }

        /// <summary>
        /// 获取或设置是否自动换行
        /// </summary>
        bool WrapText { get; set; }

        /// <summary>
        /// 获取或设置垂直对齐方式
        /// </summary>
        ETVerticalAlignment VerticalAlignment { get; set; }

        /// <summary>
        /// 边框样式
        /// </summary>
        ETBorderStyle BorderStyle { get; set; }
        /// <summary>
        /// 边框颜色
        /// </summary>
        string BorderColor { get; set; }
    }
}
