using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTemplate.Attributes;
using NPOI.SS.Formula.Functions;

namespace ExcelTemplate.Test.Model
{
    internal class ListModel
    {
        [Position("C5")]
        public List<ListItem> Children { get; set; }


        internal class ListItem
        {
            [Title("列1", "C5")]
            public int Field_1 { get; set; }

            [Title("列2", "D5")]
            public int Field_2 { get; set; }

            [Title("列3", "E5")]
            public string Field_3 { get; set; }

            [Title("列4", "F5")]
            public DateTime Field_4 { get; set; }
        }
    }
}
