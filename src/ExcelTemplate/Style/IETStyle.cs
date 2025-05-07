using System;
using NPOI.SS.UserModel;

namespace ExcelTemplate.Style
{
    /// <summary>
    /// 全量样式，大部分字段跟NPOI的样式对应
    /// </summary>
    public interface IETStyle : ICloneable
    {
        /// <summary>
        /// 单元格是否自动缩小以适应文本（当文本过长时）
        /// </summary>
        bool ShrinkToFit { get; set; }

        /// <summary>
        /// 获取或设置数据格式
        /// </summary>
        string DataFormat { get; set; }

        /// <summary>
        /// 设置字体样式
        /// </summary>
        ETFont Font { get; set; }

        /// <summary>
        /// 获取或设置单元格是否隐藏
        /// </summary>
        bool IsHidden { get; set; }

        /// <summary>
        /// 获取或设置单元格是否锁定
        /// </summary>
        bool IsLocked { get; set; }

        /// <summary>
        /// 设置或获取是否添加"引用前缀"或"123前缀"
        /// （用于告诉Excel将看似数字或公式的内容视为文本）
        /// 启用此功能类似于在Excel单元格值前添加单引号
        /// </summary>
        bool IsQuotePrefixed { get; set; }

        /// <summary>
        /// 获取或设置水平对齐方式
        /// </summary>
        NPOI.SS.UserModel.HorizontalAlignment Alignment { get; set; }

        /// <summary>
        /// 获取或设置是否自动换行
        /// </summary>
        bool WrapText { get; set; }

        /// <summary>
        /// 获取或设置垂直对齐方式
        /// </summary>
        NPOI.SS.UserModel.VerticalAlignment VerticalAlignment { get; set; }

        /// <summary>
        /// 获取或设置文本旋转角度
        /// 注意：HSSF使用-90到90度，XSSF使用0到180度
        /// 该方法会自动在这两种范围间转换
        /// </summary>
        short Rotation { get; set; }

        /// <summary>
        /// 获取或设置文本缩进量（空格数）
        /// </summary>
        short Indention { get; set; }

        /// <summary>
        /// 获取或设置左边框样式
        /// </summary>
        BorderStyle BorderLeft { get; set; }

        /// <summary>
        /// 获取或设置右边框样式
        /// </summary>
        BorderStyle BorderRight { get; set; }

        /// <summary>
        /// 获取或设置上边框样式
        /// </summary>
        BorderStyle BorderTop { get; set; }

        /// <summary>
        /// 获取或设置下边框样式
        /// </summary>
        BorderStyle BorderBottom { get; set; }

        /// <summary>
        /// 获取或设置左边框颜色索引
        /// </summary>
        string LeftBorderColor { get; set; }

        /// <summary>
        /// 获取或设置右边框颜色索引
        /// </summary>
        string RightBorderColor { get; set; }

        /// <summary>
        /// 获取或设置上边框颜色索引
        /// </summary>
        string TopBorderColor { get; set; }

        /// <summary>
        /// 获取或设置下边框颜色索引
        /// </summary>
        string BottomBorderColor { get; set; }

        /// <summary>
        /// 获取或设置填充模式
        /// （设置为1表示使用前景色填充）
        /// </summary>
        FillPattern FillPattern { get; set; }

        /// <summary>
        /// 获取或设置背景填充颜色索引
        /// </summary>
        string FillBackgroundColor { get; set; }

        /// <summary>
        /// 获取或设置前景填充颜色索引
        /// </summary>
        string FillForegroundColor { get; set; }

        /// <summary>
        /// 获取或设置对角线边框颜色
        /// </summary>
        string BorderDiagonalColor { get; set; }

        /// <summary>
        /// 获取或设置对角线边框线型
        /// </summary>
        BorderStyle BorderDiagonalLineStyle { get; set; }

        /// <summary>
        /// 获取或设置对角线边框类型
        /// </summary>
        BorderDiagonal BorderDiagonal { get; set; }

        /// <summary>
        /// 获取对应的 ICellStyle
        /// </summary>
        /// <returns></returns>
        ICellStyle GetCellStyle(IWorkbook workbook);
    }
}
