﻿using ExcelTemplate.Attributes;

namespace ExcelTemplate.Test.Model
{
    public class FormModel
    {
        [Title("字段1：", "B1"), Position("C1")]
        public int Field_1 { get; set; }

        [Title("字段2：", "B2"), Position("C2")]
        public int Field_2 { get; set; }

        [Title("测试标题1", "E7"), Position("F7")]
        public string Field_3 { get; set; }

        [Title("测试标题2", "C8"), Position("D8")]
        public DateTime Field_4 { get; set; }
    }
}
