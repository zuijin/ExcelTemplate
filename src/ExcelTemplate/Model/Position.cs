using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using ExcelTemplate.Helper;
using NPOI.SS.Formula.Functions;

namespace ExcelTemplate.Model
{
    public partial class Position : ICloneable
    {
        const string LETTER_FORMAT = "^([A-Z]+)([0-9]+)$";

        int _row = 0;
        int _col = 0;
        string _letter = "A1";

        public Position(int row, int col)
        {
            this.Row = row;
            this.Col = col;
        }

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
        private void SetLetter(string letter)
        {
            if (!IsPositionLetter(letter)) throw new Exception("letter格式错误");

            _letter = letter;
            _col = LetterHelper.ParseCol(letter);
            _row = LetterHelper.ParseRow(letter);
        }

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
