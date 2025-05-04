using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTemplate.Attributes;
using NPOI.SS.Formula.Functions;

namespace ExcelTemplate.Test.Model
{
    internal class MergeHeaderListModel
    {
        [Position("B6")]
        public List<ListItem> Children { get; set; }


        internal class ListItem
        {
            [Merge("第一级", "第二级")]
            [Title("列1", "B8")]
            public int Field_1 { get; set; }

            [Merge("第一级", "第二级")]
            [Title("列2：", "C8")]
            public int Field_2 { get; set; }

            [Merge("第一级")]
            [Title("列3", "D7")]
            public string Field_3 { get; set; }

            [Title("列4", "E6")]
            public DateTime Field_4 { get; set; }
        }
    }
}
