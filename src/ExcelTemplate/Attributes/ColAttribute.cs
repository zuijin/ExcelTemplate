using System;
using ExcelTemplate.Model;

namespace ExcelTemplate.Attributes
{
    /// <summary>
    /// Table 列标记
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ColAttribute : Attribute
    {

        public ColAttribute(string headerText, string position)
        {
            this.HeaderText = headerText;
            this.Position = position;
        }

        public ColAttribute(string headerText, int row, int col)
        {
            this.HeaderText = headerText;
            this.Position = (row, col);
        }

        public string HeaderText { get; set; }

        public Position Position { get; set; }
    }
}
