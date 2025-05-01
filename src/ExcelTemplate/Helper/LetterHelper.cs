using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.Formula.Functions;
using System.Text.RegularExpressions;

namespace ExcelTemplate.Helper
{
    public static class LetterHelper
    {
        const string LETTER_FORMAT = "^([A-Z]+)([0-9]+)$";

        /// <summary>
        /// 重置 Row 和 Col
        /// </summary>
        /// <param name="letter"></param>
        /// <exception cref="Exception"></exception>
        public static int ParseRow(string letter)
        {
            var match = Regex.Match(letter, LETTER_FORMAT);
            if (match.Groups.Count != 3)
            {
                throw new Exception("Cell位置错误，请输入正确的 Letter 格式");
            }

            return int.Parse(match.Groups[2].Value) - 1;
        }

        public static int ParseCol(string letter)
        {
            var match = Regex.Match(letter, LETTER_FORMAT);
            if (match.Groups.Count != 3)
            {
                throw new Exception("Cell位置错误，请输入正确的 Letter 格式");
            }

            return ConvertFromBase26(match.Groups[1].Value) - 1;
        }

        public static string GetLetter(int row, int col)
        {
            var colLetter = ConvertToBase26(col + 1);
            var rowLetter = row + 1;

            return $"{colLetter}{rowLetter}";
        }

        public static string ConvertToBase26(int number)
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

        public static int ConvertFromBase26(string letter)
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

                result = result * 26 + c - 'A' + 1;
            }

            return result;
        }
    }
}
