using System;

namespace ExcelTemplate.Model
{
    /// <summary>
    /// 单元格异常
    /// </summary>
    public class CellException : Exception
    {
        /// <summary>
        /// 
        /// </summary>
        public CellException(int row, int column, string message) : base(message)
        {
            Row = row;
            Column = column;
        }

        public CellException(int row, int column, string message, Exception inner) : base(message, inner)
        {
            Row = row;
            Column = column;
        }

        /// <summary>
        /// 行号
        /// </summary>
        public int Row { get; set; }

        /// <summary>
        /// 列号
        /// </summary>
        public int Column { get; set; }

    }
}
