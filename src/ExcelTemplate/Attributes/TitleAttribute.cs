using System;
using ExcelTemplate.Model;

namespace ExcelTemplate.Attributes
{
    /// <summary>
    /// 文本标题标记
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public class TitleAttribute : Attribute
    {
        public TitleAttribute(string title, string position) : this(title, position, null)
        {

        }

        public TitleAttribute(string title, string position, string mergeTo)
        {
            this.Title = title;
            this.Position = position;
            this.MergeTo = mergeTo;
        }

        public TitleAttribute(string title, int row, int col)
        {
            this.Title = title;
            this.Position = (row, col);
        }

        public TitleAttribute(string title, int row, int col, int mergeToRow, int mergeToCol)
        {
            this.Title = title;
            this.Position = (row, col);
            this.MergeTo = (mergeToCol, mergeToRow);
        }

        public string Title { get; set; }

        public Position Position { get; set; }

        public Position MergeTo { get; set; }
    }
}
