using ExcelTemplate.Model;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTemplate.Hint
{
    public class CellHintMessage
    {
        public CellHintMessage(int row, int col, string message)
        {
            this.Position = new Position(row, col);
            this.Message = message;
        }


        public CellHintMessage(string letter, string message)
        {
            this.Position = letter;
            this.Message = message;
        }

        public CellHintMessage(Position position, string message)
        {
            this.Position = position;
            this.Message = message;
        }

        /// <summary>
        /// 单元格位置
        /// </summary>
        public Position Position { get; set; }

        /// <summary>
        /// 提示信息
        /// </summary>
        public string Message { get; set; }
    }
}
