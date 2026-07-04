using System;
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
            letter = letter?.ToUpper();
            var match = Regex.Match(letter, LETTER_FORMAT);
            if (match.Groups.Count != 3)
            {
                throw new Exception("Cell位置错误，请输入正确的 Letter 格式");
            }

            return int.Parse(match.Groups[2].Value) - 1;
        }

        /// <summary>
        /// 解析列
        /// </summary>
        /// <param name="letter">字母</param>
        /// <returns>列索引</returns>
        public static int ParseCol(string letter)
        {
            letter = letter?.ToUpper();
            var match = Regex.Match(letter, LETTER_FORMAT);
            if (match.Groups.Count != 3)
            {
                throw new Exception("Cell位置错误，请输入正确的 Letter 格式");
            }

            return ConvertFromBase26(match.Groups[1].Value) - 1;
        }

        /// <summary>
        /// 获取行和列的字母表示
        /// </summary>
        /// <param name="row">行索引</param>
        /// <param name="col">列索引</param>
        /// <returns>字母表示</returns>
        public static string GetLetter(int row, int col)
        {
            var colLetter = ConvertToBase26(col + 1);
            var rowLetter = row + 1;

            return $"{colLetter}{rowLetter}";
        }

        /// <summary>
        /// 转换为26进制表示
        /// </summary>
        /// <param name="number">数字</param>
        /// <returns>26进制表示的字符串</returns>
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

        /// <summary>
        /// 从26进制表示转换回数字
        /// </summary>
        /// <param name="letter">26进制表示的字符串</param>
        /// <returns>数字</returns>
        public static int ConvertFromBase26(string letter)
        {
            if (string.IsNullOrEmpty(letter))
            {
                throw new ArgumentException("letter不能为空");
            }

            letter = letter.ToUpper();
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
