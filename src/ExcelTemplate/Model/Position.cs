using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using ExcelTemplate.Helper;
using NPOI.SS.Formula.Functions;

namespace ExcelTemplate.Model
{
    public class Position
    {
        const string LETTER_FORMAT = "^([A-Z]+)([0-9]+)$";

        public Position(int row, int col)
        {
            if (row < 0)
                throw new Exception("row不能小于0");
            if (col < 0)
                throw new Exception("col不能小于0");

            this.Row = row;
            this.Col = col;
            this.Letter = LetterHelper.GetLetter(row, col);
        }

        public Position(string letter)
        {
            if (!IsPositionLetter(letter))
                throw new Exception("letter格式错误");

            this.Letter = letter;
            this.Col = LetterHelper.ParseCol(letter);
            this.Row = LetterHelper.ParseRow(letter);
        }

        /// <summary>
        /// 行下标，从 0 开始
        /// </summary>
        public int Row { get; private set; }
        /// <summary>
        /// 列下标，从 0 开始
        /// </summary>
        public int Col { get; private set; }
        /// <summary>
        /// 字符表示的单元格位置，例如 A1 表示第一行第一列
        /// </summary>
        public string Letter { get; private set; } = "A1";

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

        public static implicit operator Position?(string? v)
        {
            if (string.IsNullOrWhiteSpace(v))
            {
                return null;
            }

            return new Position(v);
        }
    }

}
