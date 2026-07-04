using ExcelTemplate.Model;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTemplate.Hint
{
    public class CellHintMessage
    {
        /// <summary>
        /// 实例化单元格提示信息
        /// </summary>
        /// <param name="row">行索引</param>
        /// <param name="col">列索引</param>
        /// <param name="message">提示信息</param>
        public CellHintMessage(int row, int col, string message)
        {
            this.Position = new Position(row, col);
            this.Message = message;
        }

        /// <summary>
        /// 实例化单元格提示信息
        /// </summary>
        /// <param name="letter">字母表示的位置</param>
        /// <param name="message">提示信息</param>
        public CellHintMessage(string letter, string message)
        {
            this.Position = letter;
            this.Message = message;
        }

        /// <summary>
        /// 实例化单元格提示信息
        /// </summary>
        /// <param name="position">位置对象</param>
        /// <param name="message">提示信息</param>
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
