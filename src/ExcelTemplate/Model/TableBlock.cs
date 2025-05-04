using System;
using System.Collections.Generic;

namespace ExcelTemplate.Model
{
    public class TableBlock : BlockBase, ICloneable
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


        public override object Clone()
        {
            var obj = (TableBlock)base.Clone();
            if (this.Header != null)
            {
                obj.Header = new List<TableHeaderBlock>();
                foreach (var item in this.Header)
                {
                    obj.Header.Add((TableHeaderBlock)item.Clone());
                }
            }

            if (this.Body != null)
            {
                obj.Body = new List<TableBodyBlock>();
                foreach (var item in this.Body)
                {
                    obj.Body.Add((TableBodyBlock)item.Clone());
                }
            }

            return obj;
        }
    }
}
