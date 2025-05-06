using System;

namespace ExcelTemplate.Attributes
{
    /// <summary>
    /// Table 列标记
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ColAttribute : Attribute
    {

        public ColAttribute(string headerText, int colIndex)
        {
            this.HeaderText = headerText;
            this.ColIndex = colIndex;
        }

        public ColAttribute(string headerText, int row, int col)
        {
            this.HeaderText = headerText;
        }

        /// <summary>
        /// 表头
        /// </summary>
        public string HeaderText { get; set; }

        /// <summary>
        /// 列顺序，从0开始递增
        /// </summary>
        public int ColIndex { get; set; }
    }
}
