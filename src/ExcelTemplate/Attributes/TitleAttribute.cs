using System;

namespace ExcelTemplate.Attributes
{
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

        public string Title { get; set; }

        public string Position { get; set; }

        public string MergeTo { get; set; }
    }
}
