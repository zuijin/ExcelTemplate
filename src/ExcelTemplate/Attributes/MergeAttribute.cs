using System;

namespace ExcelTemplate.Attributes
{
    public class MergeAttribute : Attribute
    {
        public MergeAttribute(params string[] titles)
        {
            Titles = titles;
        }

        public string[] Titles { get; set; }
    }
}
