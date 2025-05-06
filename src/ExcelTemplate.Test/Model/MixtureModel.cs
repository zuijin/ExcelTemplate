using ExcelTemplate.Attributes;

namespace ExcelTemplate.Test.Model
{
    internal class MixtureModel
    {
        [Title("姓名：", "B2")]
        [Position("C2")]
        public string StudentName { get; set; }

        [Title("性别：", "D2")]
        [Position("E2")]
        public string Sex { get; set; }

        [Title("生日：", "B3")]
        [Position("C3")]
        public DateTime BirthDate { get; set; }

        [Title("第一次模拟考试", "B5", "E5")]
        [Position("B6")]
        public List<ScoreItem> Scores_1st { get; set; }

        [Title("总分：", "D10")]
        [Position("E10")]
        public float TotalScore_1st { get; set; }

        [Title("总排名：", "D11")]
        [Position("E11")]
        public int TotalRanking_1st { get; set; }

        [Title("第二次模拟考试", "B13", "E13")]
        [Position("B14")]
        public List<ScoreItem> Scores_2nd { get; set; }

        [Title("总分：", "D18")]
        [Position("E18")]
        public float TotalScore_2nd { get; set; }

        [Title("总排名：", "D19")]
        [Position("E19")]
        public int TotalRanking_2nd { get; set; }


        internal class ScoreItem
        {
            [Col("科目", 0)]
            public string SubjectName { get; set; }

            [Merge("成绩")]
            [Col("分数", 1)]
            public float Score { get; set; }

            [Merge("成绩")]
            [Col("排名", 2)]
            public int Ranking { get; set; }

            [Col("考试时间", 3)]
            public DateTime ExamTime { get; set; }
        }
    }
}
