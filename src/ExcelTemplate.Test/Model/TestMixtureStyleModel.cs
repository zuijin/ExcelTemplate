using ExcelTemplate.Attributes;
using ExcelTemplate.Style;

namespace ExcelTemplate.Test.Model
{
    [StyleDic(Key = "field_title", FontHeightInPoints = 11, IsBold = true, HorizontalAlignment = ETHorizontalAlignment.Left, VerticalAlignment = ETVerticalAlignment.Center, ShrinkToFit = true)]
    [StyleDic(Key = "field_value", FontHeightInPoints = 11, IsBold = false, HorizontalAlignment = ETHorizontalAlignment.Left, VerticalAlignment = ETVerticalAlignment.Center, ShrinkToFit = true)]
    [StyleDic(Key = "field_value_number", FontHeightInPoints = 11, IsBold = false, HorizontalAlignment = ETHorizontalAlignment.Right, VerticalAlignment = ETVerticalAlignment.Center, ShrinkToFit = true)]
    [StyleDic(Key = "field_value_date", FontHeightInPoints = 11, IsBold = false, HorizontalAlignment = ETHorizontalAlignment.Left, VerticalAlignment = ETVerticalAlignment.Center, ShrinkToFit = true, DataFormat = "yyyy/m/d")]
    [StyleDic(Key = "table_head", FontHeightInPoints = 11, IsBold = true, HorizontalAlignment = ETHorizontalAlignment.Center, VerticalAlignment = ETVerticalAlignment.Center, ShrinkToFit = true, BgColor = "EE822F", BorderStyle = ETBorderStyle.Thin)]
    [StyleDic(Key = "table_body", FontHeightInPoints = 11, IsBold = false, HorizontalAlignment = ETHorizontalAlignment.Left, VerticalAlignment = ETVerticalAlignment.Center, ShrinkToFit = true, BorderStyle = ETBorderStyle.Thin)]
    [StyleDic(Key = "table_body_date", FontHeightInPoints = 11, IsBold = false, HorizontalAlignment = ETHorizontalAlignment.Right, VerticalAlignment = ETVerticalAlignment.Center, ShrinkToFit = true, BorderStyle = ETBorderStyle.Thin, DataFormat = "yyyy/m/d")]
    [StyleDic(Key = "table_body_number", FontHeightInPoints = 11, IsBold = false, HorizontalAlignment = ETHorizontalAlignment.Right, VerticalAlignment = ETVerticalAlignment.Center, ShrinkToFit = true, BorderStyle = ETBorderStyle.Thin)]
    internal class TestMixtureStyleModel
    {
        [Title("姓名：", "B2", Style = "field_title")]
        [Position("C2", Style = "field_value")]
        public string StudentName { get; set; }

        [Title("性别：", "D2", Style = "field_title")]
        [Position("E2", Style = "field_value")]
        public string Sex { get; set; }

        [Title("生日：", "B3", Style = "field_title")]
        [Position("C3", Style = "field_value_date")]
        public DateTime BirthDate { get; set; }

        [Title("第一次模拟考试", "B5", "E5", Style = "table_head")]
        [Position("B6", Style = "table_head")]
        public List<ScoreItem> Scores_1st { get; set; }

        [Title("总分：", "D10", Style = "field_title")]
        [Position("E10", Style = "field_value_number")]
        public float TotalScore_1st { get; set; }

        [Title("总排名：", "D11", Style = "field_title")]
        [Position("E11", Style = "field_value_number")]
        public int TotalRanking_1st { get; set; }

        [Title("第二次模拟考试", "B13", "E13", Style = "table_head")]
        [Position("B14", Style = "table_head")]
        public List<ScoreItem> Scores_2nd { get; set; }

        [Title("总分：", "D18", Style = "field_title")]
        [Position("E18", Style = "field_value_number")]
        public float TotalScore_2nd { get; set; }

        [Title("总排名：", "D19", Style = "field_title")]
        [Position("E19", Style = "field_value_number")]
        public int TotalRanking_2nd { get; set; }


        internal class ScoreItem
        {
            [Col("科目", 0, Style = "table_body")]
            public string SubjectName { get; set; }

            [Merge("成绩")]
            [Col("分数", 1, Style = "table_body_number")]
            public float Score { get; set; }

            [Merge("成绩")]
            [Col("排名", 2, Style = "table_body_number")]
            public int Ranking { get; set; }

            [Col("考试时间", 3, Style = "table_body_date")]
            public DateTime ExamTime { get; set; }
        }
    }
}
