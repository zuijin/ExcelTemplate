using System;
using ExcelTemplate.Model;

namespace ExcelTemplate.Exceptions
{
    /// <summary>
    /// 单元格异常
    /// </summary>
    public class CellException : Exception
    {
        /// <summary>
        /// 实例化单元格异常
        /// </summary>
        /// <param name="row">行索引</param>
        /// <param name="col">列索引</param>
        /// <param name="message">异常信息</param>
        public CellException(int row, int col, string message) : base(message)
        {
            Position = new Position(row, col);
        }

        /// <summary>
        /// 实例化单元格异常
        /// </summary>
        /// <param name="row">行索引</param>
        /// <param name="col">列索引</param>
        /// <param name="message">异常信息</param>
        /// <param name="inner">内部异常</param>
        public CellException(int row, int col, string message, Exception inner) : base(message, inner)
        {
            Position = new Position(row, col);
        }

        /// <summary>
        /// 实例化单元格异常
        /// </summary>
        /// <param name="letter">字母表示的单元格位置</param>
        /// <param name="message">异常信息</param>
        public CellException(string letter, string message) : base(message)
        {
            Position = letter;
        }

        /// <summary>
        /// 实例化单元格异常
        /// </summary>
        /// <param name="letter">字母表示的单元格位置</param>
        /// <param name="message">异常信息</param>
        /// <param name="inner">内部异常</param>
        public CellException(string letter, string message, Exception inner) : base(message, inner)
        {
            Position = letter;
        }

        /// <summary>
        /// 异常发生的单元格位置
        /// </summary>
        public Position Position { get; set; }
    }
}
