using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using NPOI.SS.Formula.Functions;

namespace ExcelTemplate
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

            Row = row;
            Col = col;
            ResetLetter(row, col);
        }

        public Position(string letter)
        {
            if (!Regex.IsMatch(letter, LETTER_FORMAT))
                throw new Exception("letter格式错误");

            Letter = letter;
            ResetRowCol(letter);
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

        public Position GetNextRow()
        {
            return new Position(this.Row + 1, this.Col);
        }

        public Position GetNextCol()
        {
            return new Position(this.Row, this.Col + 1);
        }

        /// <summary>
        /// 重置 Row 和 Col
        /// </summary>
        /// <param name="letter"></param>
        /// <exception cref="Exception"></exception>
        private void ResetRowCol(string letter)
        {
            var match = Regex.Match(letter, LETTER_FORMAT);
            if (match.Groups.Count != 2)
            {
                throw new Exception("Cell位置错误，请输入正确的 Letter 格式");
            }

            this.Col = ConvertFromBase26(match.Groups[1].Value);
            this.Row = int.Parse(match.Groups[2].Value) - 1;
        }

        private void ResetLetter(int row, int col)
        {
            var colLetter = ConvertToBase26(col + 1);
            var rowLetter = row + 1;

            this.Letter = $"{colLetter}{rowLetter}";
        }


        private string ConvertToBase26(int number)
        {
            if (number <= 0)
            {
                throw new ArgumentException("number 不能小于1");
            }

            string result = "";
            while (number > 0)
            {
                number--; // 因为A对应1而不是0
                int remainder = number % 26;
                char digit = (char)('A' + remainder);
                result = digit + result;
                number /= 26;
            }

            return result;
        }

        private int ConvertFromBase26(string letter)
        {
            if (string.IsNullOrEmpty(letter))
            {
                throw new ArgumentException("letter不能为空");
            }

            int result = 0;
            for (int i = 0; i < letter.Length; i++)
            {
                char c = letter[i];
                if (c < 'A' || c > 'Z')
                {
                    throw new ArgumentException("输入字符只能为 A-Z.");
                }

                result = result * 26 + (c - 'A' + 1);
            }

            return result;
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
