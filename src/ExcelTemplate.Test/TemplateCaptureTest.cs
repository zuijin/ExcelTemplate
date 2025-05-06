using ExcelTemplate.Test.Model;

namespace ExcelTemplate.Test
{
    public class TemplateCaptureTest
    {
        /// <summary>
        /// 读表单
        /// </summary>
        [Fact]
        public void TestReadForm()
        {
            var filePath = "Files/Form.xlsx";
            var file = File.Open(filePath, FileMode.Open);
            var template = TemplateCapture.Create(typeof(FormModel));

            dynamic data = template.Capture<FormModel>(file);
            Assert.Equal(data.Field_1, 123);
            Assert.Equal(data.Field_2, 456);
            Assert.Equal(data.Field_3, "aabcc");
            Assert.Equal(data.Field_4, DateTime.Parse("2025/1/2"));
        }

        /// <summary>
        /// 读取列表
        /// </summary>
        [Fact]
        public void TestReadList()
        {
            var filePath = "Files/List.xlsx";
            var file = File.Open(filePath, FileMode.Open);
            var template = TemplateCapture.Create(typeof(ListModel));

            dynamic data = template.Capture<ListModel>(file);
            Assert.Equal(data.Children.Count, 7);

            object[][] tmp =
            [
                [123,456,"eeee",DateTime.Parse("2000/5/6")],
                [124,457,"aa",DateTime.Parse("2000/5/7")],
                [125,458,"bb",DateTime.Parse("2000/5/8")],
                [126,459,"cc",DateTime.Parse("2000/5/9")],
                [127,460,"dd",DateTime.Parse("2000/5/10")],
                [128,461,"eeee",DateTime.Parse("2000/5/11")],
                [129,462,"ff",DateTime.Parse("2000/5/12")],
            ];

            for (int i = 0; i < data.Children.Count; i++)
            {
                var item = data.Children[i];
                Assert.Equal(item.Field_1, tmp[i][0]);
                Assert.Equal(item.Field_2, tmp[i][1]);
                Assert.Equal(item.Field_3, tmp[i][2]);
                Assert.Equal(item.Field_4, tmp[i][3]);
            }
        }

        /// <summary>
        /// 读取包含合并表头的列表
        /// </summary>
        [Fact]
        public void TestReadMergeHeaderList()
        {
            var filePath = "Files/MergeHeaderList.xlsx";
            var file = File.Open(filePath, FileMode.Open);
            var template = TemplateCapture.Create(typeof(MergeHeaderListModel));

            dynamic data = template.Capture<MergeHeaderListModel>(file);
            Assert.Equal(data.Children.Count, 7);

            object[][] tmp =
            [
                [123,456,"1111",DateTime.Parse("2000/5/6")],
                [124,457,"aa",DateTime.Parse("2000/5/7")],
                [125,458,"bb",DateTime.Parse("2000/5/8")],
                [126,459,"cc",DateTime.Parse("2000/5/9")],
                [127,460,"dd",DateTime.Parse("2000/5/10")],
                [128,461,"eeee",DateTime.Parse("2000/5/11")],
                [129,462,"ff",DateTime.Parse("2000/5/12")],
            ];

            for (int i = 0; i < data.Children.Count; i++)
            {
                var item = data.Children[i];
                Assert.Equal(item.Field_1, tmp[i][0]);
                Assert.Equal(item.Field_2, tmp[i][1]);
                Assert.Equal(item.Field_3, tmp[i][2]);
                Assert.Equal(item.Field_4, tmp[i][3]);
            }

        }

        /// <summary>
        /// 读取混排
        /// </summary>
        [Fact]
        public void TestReadMixture()
        {
            var filePath = "Files/Mixture.xlsx";
            var file = File.Open(filePath, FileMode.Open);
            var template = TemplateCapture.Create(typeof(MixtureModel));

            var data = template.Capture<MixtureModel>(file);
            Assert.Equal("张三", data.StudentName);
            Assert.Equal("男", data.Sex);
            Assert.Equal(DateTime.Parse("2025/2/3"), data.BirthDate);
            Assert.Equal(936f, data.TotalScore_1st);
            Assert.Equal(7, data.TotalRanking_1st);
            Assert.Equal(972f, data.TotalScore_2nd);
            Assert.Equal(5, data.TotalRanking_2nd);

            object[][] tmp =
            [
                ["语文", 100f, 2, DateTime.Parse("2000/5/6")],
                ["数学", 101f, 3, DateTime.Parse("2000/5/7")],
                ["英语", 102f, 4, DateTime.Parse("2000/5/8")],
                ["化学", 103f, 5, DateTime.Parse("2000/5/9")],
                ["物理", 104f, 6, DateTime.Parse("2000/5/10")],
                ["生物", 105f, 7, DateTime.Parse("2000/5/11")],
                ["地理", 106f, 8, DateTime.Parse("2000/5/12")],
                ["政治", 107f, 9, DateTime.Parse("2000/5/13")],
                ["历史", 108f, 10,DateTime.Parse("2000/5/14")],
            ];

            for (int i = 0; i < data.Scores_1st.Count; i++)
            {
                var item = data.Scores_1st[i];
                Assert.Equal(item.SubjectName, tmp[i][0]);
                Assert.Equal(item.Score, tmp[i][1]);
                Assert.Equal(item.Ranking, tmp[i][2]);
                Assert.Equal(item.ExamTime, tmp[i][3]);
            }

            object[][] tmp2 =
            [
                ["语文", 104f, 2, DateTime.Parse("2000/6/6")],
                ["数学", 105f, 5, DateTime.Parse("2000/6/7")],
                ["英语", 106f, 4, DateTime.Parse("2000/6/8")],
                ["化学", 107f, 1, DateTime.Parse("2000/6/9")],
                ["物理", 108f, 6, DateTime.Parse("2000/6/10")],
                ["生物", 109f, 7, DateTime.Parse("2000/6/11")],
                ["地理", 110f, 7, DateTime.Parse("2000/6/12")],
                ["政治", 111f, 9, DateTime.Parse("2000/6/13")],
                ["历史", 112f, 6, DateTime.Parse("2000/6/14")],
            ];

            for (int i = 0; i < data.Scores_2nd.Count; i++)
            {
                var item = data.Scores_2nd[i];
                Assert.Equal(item.SubjectName, tmp2[i][0]);
                Assert.Equal(item.Score, tmp2[i][1]);
                Assert.Equal(item.Ranking, tmp2[i][2]);
                Assert.Equal(item.ExamTime, tmp2[i][3]);
            }
        }
    }
}