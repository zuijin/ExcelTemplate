using System;
using ExcelTemplate.Model;

namespace ExcelTemplate.Attributes
{
    /// <summary>
    /// 单元格位置标记
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class PositionAttribute : Attribute
    {
        public PositionAttribute(string position) : this(position, null)
        {
        }

        public PositionAttribute(string position, string mergeTo)
        {
            this.Position = position;
            this.MergeTo = mergeTo;
        }

        public PositionAttribute(int row, int col)
        {
            this.Position = (row, col);
        }

        public PositionAttribute(int row, int col, int mergeToRow, int mergeToCol)
        {
            this.Position = (row, col);
            this.MergeTo = (mergeToCol, mergeToRow);
        }

        public Position Position { get; private set; }

        public Position MergeTo { get; private set; }
    }
}
