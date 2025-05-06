using ExcelTemplate.Attributes;

namespace ExcelTemplate.Test.Model
{
    internal class MergeHeaderListModel
    {
        [Position("B6")]
        public List<ListItem> Children { get; set; }


        internal class ListItem
        {
            [Merge("第一级", "第二级")]
            [Col("列1", 0)]
            public int Field_1 { get; set; }

            [Merge("第一级", "第二级")]
            [Col("列2：", 1)]
            public int Field_2 { get; set; }

            [Merge("第一级")]
            [Col("列3", 2)]
            public string Field_3 { get; set; }

            [Col("列4", 3)]
            public DateTime Field_4 { get; set; }
        }
    }
}
