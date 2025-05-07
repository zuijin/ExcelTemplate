using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTemplate.Style
{
    public enum HorizontalAlignment
    {
        /// <summary>
        /// 常规对齐（默认对齐方式）
        /// 文本数据左对齐，数字、日期和时间右对齐
        /// 布尔类型居中对齐
        /// 注意：更改对齐方式不会改变数据类型
        /// </summary>
        General = 0,

        /// <summary>
        /// 左对齐（即使在从右到左模式中）
        /// 内容与单元格左边缘对齐
        /// 如果指定了缩进量，内容将从左边缩进指定数量的字符空格
        /// 字符空格基于工作簿的默认字体和字号
        /// </summary>
        Left = 1,

        /// <summary>
        /// 居中对齐
        /// 文本在单元格内水平居中显示
        /// </summary>
        Center = 2,

        /// <summary>
        /// 右对齐（即使在从右到左模式中）
        /// 内容与单元格右边缘对齐
        /// </summary>
        Right = 3,
    }
}
