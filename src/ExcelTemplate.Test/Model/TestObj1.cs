using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTemplate.Attributes;

namespace ExcelTemplate.Test.Model
{
    public class TestObj1
    {
        [Title("测试标题", "A2"), Value( "B2")]
        public int Id { get; set; }

        [Title("测试标题2", "A3"), Value("B3")]
        public int Id2 { get; set; }
    }
}
