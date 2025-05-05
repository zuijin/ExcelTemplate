using System;

namespace ExcelTemplate.Attributes
{
    public class MergeAttribute : Attribute
    {
        public MergeAttribute(params string[] titles)
        {
            if (titles == null || titles.Length == 0)
            {
                throw new Exception($"{nameof(MergeAttribute)}必须指定{nameof(titles)}参数");
            }

            Titles = titles;
        }

        public string[] Titles { get; private set; }
    }
}
