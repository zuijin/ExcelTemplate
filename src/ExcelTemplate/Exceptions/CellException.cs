using System;
using ExcelTemplate.Model;

namespace ExcelTemplate.Exceptions
{
    /// <summary>
    /// 单元格异常
    /// </summary>
    public class CellException : Exception
    {
        public CellException(int row, int col, string message) : base(message)
        {
            Position = new Position(row, col);
        }

        public CellException(int row, int col, string message, Exception inner) : base(message, inner)
        {
            Position = new Position(row, col);
        }

        public CellException(string letter, string message) : base(message)
        {
            Position = letter;
        }

        public CellException(string letter, string message, Exception inner) : base(message, inner)
        {
            Position = letter;
        }

        public Position Position { get; set; }
    }
}
