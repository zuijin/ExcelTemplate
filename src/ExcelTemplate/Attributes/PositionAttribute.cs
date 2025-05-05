using System;

namespace ExcelTemplate.Attributes
{
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

        public string Position { get; private set; }

        public string MergeTo { get; private set; }
    }
}
