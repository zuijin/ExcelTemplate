using System;

namespace ExcelTemplate.Hint
{
    public class RowHintMessage : ICloneable
    {
        /// <summary>
        /// 实例化行提示信息
        /// </summary>
        /// <param name="row">行索引</param>
        /// <param name="message">提示信息</param>
        public RowHintMessage(int row, string message)
        {
            this.Row = row;
            this.Message = message;
        }

        /// <summary>
        /// 行号
        /// </summary>
        public int Row { get; set; }

        /// <summary>
        /// 提示信息
        /// </summary>
        public string Message { get; set; }

        public object Clone()
        {
            var obj = (CellHintMessage)this.MemberwiseClone();
            return obj;
        }
    }
}
