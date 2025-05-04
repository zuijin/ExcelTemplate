using System;

namespace ExcelTemplate.Model
{
    /// <summary>
    /// 单元格异常
    /// </summary>
    public class CellException : Exception
    {
        public CellException(int row, int col, string message) : base(message)
        {
            this.Position = new Position(row, col);
        }

        public CellException(int row, int col, string message, Exception) : base(message, inner)
        {
            this.Position = new Position(row, col);
        }

        public CellException(string letter, string message) : base(message)
        {
            this.Position = letter;
        }

        public CellException(string letter, string message, Exception) : base(message, inner)
        {
            this.Position = letter;
        }

        public Position Position { get; set; }
    }
}
