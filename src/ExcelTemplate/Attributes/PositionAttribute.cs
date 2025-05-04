using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTemplate.Attributes
{
    public class PositionAttribute : Attribute
    {
        public PositionAttribute(string position) : this(position, null)
        {
        }

        public PositionAttribute(string position, string? mergeTo)
        {
            this.Position = position;
            this.MergeTo = mergeTo;
        }

        public string Position { get; set; }

        public string? MergeTo { get; set; }
    }
}
