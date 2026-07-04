using System;
using System.Text.RegularExpressions;
using ExcelTemplate.Helper;

namespace ExcelTemplate.Model
{
    public partial class Position : ICloneable
    {
        const string LETTER_FORMAT = "^([a-zA-Z]+)([0-9]+)$";

        int _row = 0;
        int _col = 0;
        string _letter = "A1";

        /// <summary>
        /// 实例化位置对象
        /// </summary>
        /// <param name="row">行索引</param>
        /// <param name="col">列索引</param>
        public Position(int row, int col)
        {
            this.Row = row;
            this.Col = col;
        }

        /// <summary>
        /// 实例化位置对象
        /// </summary>
        /// <param name="letter">字母表示的单元格位置</param>
        public Position(string letter)
        {
            this.Letter = letter;
        }

        /// <summary>
        /// 行下标，从 0 开始
        /// </summary>
        public int Row
        {
            get => _row;
            set => SetRowCol(value, this._col);
        }

        /// <summary>
        /// 列下标，从 0 开始
        /// </summary>
        public int Col
        {
            get => _col;
            set => SetRowCol(this._row, value);
        }

        /// <summary>
        /// 字符表示的单元格位置，例如 A1 表示第一行第一列
        /// </summary>
        public string Letter
        {
            get => _letter;
            set => SetLetter(value);
        }
    }



    public partial class Position
    {
        /// <summary>
        /// 设置字母表示的位置
        /// </summary>
        /// <param name="letter">字母表示的单元格位置</param>
        private void SetLetter(string letter)
        {
            if (!IsPositionLetter(letter)) throw new Exception("letter格式错误");

            _letter = letter.ToUpper(); ;
            _col = LetterHelper.ParseCol(_letter);
            _row = LetterHelper.ParseRow(_letter);
        }

        /// <summary>
        /// 设置行和列索引
        /// </summary>
        /// <param name="row">行索引</param>
        /// <param name="col">列索引</param>
        private void SetRowCol(int row, int col)
        {
            if (row < 0) throw new Exception("row不能小于0");
            if (col < 0) throw new Exception("col不能小于0");

            _row = row;
            _col = col;
            _letter = LetterHelper.GetLetter(row, col);
        }

        /// <summary>
        /// 获取一个新的偏移位置
        /// </summary>
        /// <param name="rowOffset"></param>
        /// <param name="colOffset"></param>
        /// <returns></returns>
        public Position GetOffset(int rowOffset, int colOffset)
        {
            return new Position(Row + rowOffset, Col + colOffset);
        }

        /// <summary>
        /// 使当前位置信息进行偏移
        /// </summary>
        /// <param name="rowOffset"></param>
        /// <param name="colOffset"></param>
        public void ApplyOffset(int rowOffset = 0, int colOffset = 0)
        {
            if (rowOffset == 0 && colOffset == 0)
            {
                return;
            }

            this.Row += rowOffset;
            this.Col += colOffset;
            this.Letter = LetterHelper.GetLetter(this.Row, this.Col);
        }

        /// <summary>
        /// 判断是否符合位置格式的字符
        /// </summary>
        /// <param name="letter"></param>
        /// <returns></returns>
        public static bool IsPositionLetter(string letter)
        {
            return Regex.IsMatch(letter, LETTER_FORMAT);
        }

        /// <summary>
        /// 尝试解析字母表示的位置
        /// </summary>
        /// <param name="letter">字母表示的位置</param>
        /// <param name="pos">输出的位置对象</param>
        /// <returns>是否解析成功</returns>
        public static bool TryParse(string letter, out Position? pos)
        {
            if (!IsPositionLetter(letter))
            {
                pos = null;
                return false;
            }

            pos = new Position(letter);
            return true;
        }

        /// <summary>
        /// 转换为字符串表示
        /// </summary>
        /// <returns>字母表示的位置</returns>
        public override string ToString()
        {
            return this.Letter;
        }

        public static implicit operator Position((int, int) v)
        {
            return new Position(v.Item1, v.Item2);
        }

        public static implicit operator Position(string v)
        {
            if (string.IsNullOrWhiteSpace(v))
            {
                return null;
            }

            return new Position(v);
        }

        public object Clone()
        {
            return this.GetOffset(0, 0);
        }
    }

}
