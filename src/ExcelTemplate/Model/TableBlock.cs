using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTemplate.Model
{
    public class TableBlock : BlockBase
    {
        /// <summary>
        /// 表名
        /// </summary>
        public string TableName { get; set; }
        /// <summary>
        /// 表头
        /// </summary>
        public List<TableHeaderBlock> Header { get; set; }
        /// <summary>
        /// 表体
        /// </summary>
        public List<TableBodyBlock> Body { get; set; }
    }

    public class TableHeaderBlock : BlockBase
    {
        public string Text { get; set; }
    }

    public class TableBodyBlock : BlockBase
    {
        public string FieldPath { get; set; }
    }
}
