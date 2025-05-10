using ExcelTemplate.Test.Model;

namespace ExcelTemplate.Test
{
    [TestClass]
    public class TemplateCaptureTest
    {
        /// <summary>
        /// ����
        /// </summary>
        [TestMethod]
        public void TestReadForm()
        {
            var filePath = "Files/Form.xlsx";
            var file = File.Open(filePath, FileMode.Open);
            var template = TemplateCapture.Create(typeof(FormModel));

            dynamic data = template.Capture<FormModel>(file);
            Assert.AreEqual(data.Field_1, 123);
            Assert.AreEqual(data.Field_2, 456);
            Assert.AreEqual(data.Field_3, "aabcc");
            Assert.AreEqual(data.Field_4, DateTime.Parse("2025/1/2"));
        }

        /// <summary>
        /// ��ȡ�б�
        /// </summary>
        [TestMethod]
        public void TestReadList()
        {
            var filePath = "Files/List.xlsx";
            var file = File.Open(filePath, FileMode.Open);
            var template = TemplateCapture.Create(typeof(ListModel));

            dynamic data = template.Capture<ListModel>(file);
            Assert.AreEqual(data.Children.Count, 7);

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
                Assert.AreEqual(item.Field_1, tmp[i][0]);
                Assert.AreEqual(item.Field_2, tmp[i][1]);
                Assert.AreEqual(item.Field_3, tmp[i][2]);
                Assert.AreEqual(item.Field_4, tmp[i][3]);
            }
        }

        /// <summary>
        /// ��ȡ�����ϲ���ͷ���б�
        /// </summary>
        [TestMethod]
        public void TestReadMergeHeaderList()
        {
            var filePath = "Files/MergeHeaderList.xlsx";
            var file = File.Open(filePath, FileMode.Open);
            var template = TemplateCapture.Create(typeof(MergeHeaderListModel));

            dynamic data = template.Capture<MergeHeaderListModel>(file);
            Assert.AreEqual(data.Children.Count, 7);

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
                Assert.AreEqual(item.Field_1, tmp[i][0]);
                Assert.AreEqual(item.Field_2, tmp[i][1]);
                Assert.AreEqual(item.Field_3, tmp[i][2]);
                Assert.AreEqual(item.Field_4, tmp[i][3]);
            }

        }

        /// <summary>
        /// ��ȡ����
        /// </summary>
        [TestMethod]
        public void TestReadMixture()
        {
            var filePath = "Files/Mixture.xlsx";
            var file = File.Open(filePath, FileMode.Open);
            var template = TemplateCapture.Create(typeof(MixtureModel));

            var data = template.Capture<MixtureModel>(file);
            //Assert.AreEqual("����", data.StudentName);
            Assert.AreEqual("��", data.Sex);
            Assert.AreEqual(DateTime.Parse("2025/2/3"), data.BirthDate);
            Assert.AreEqual(936f, data.TotalScore_1st);
            Assert.AreEqual(7, data.TotalRanking_1st);
            Assert.AreEqual(972f, data.TotalScore_2nd);
            Assert.AreEqual(5, data.TotalRanking_2nd);

            object[][] tmp =
            [
                ["����", 100f, 2, DateTime.Parse("2000/5/6")],
                ["��ѧ", 101f, 3, DateTime.Parse("2000/5/7")],
                ["Ӣ��", 102f, 4, DateTime.Parse("2000/5/8")],
                ["��ѧ", 103f, 5, DateTime.Parse("2000/5/9")],
                ["����", 104f, 6, DateTime.Parse("2000/5/10")],
                ["����", 105f, 7, DateTime.Parse("2000/5/11")],
                ["����", 106f, 8, DateTime.Parse("2000/5/12")],
                ["����", 107f, 9, DateTime.Parse("2000/5/13")],
                ["��ʷ", 108f, 10,DateTime.Parse("2000/5/14")],
            ];

            for (int i = 0; i < data.Scores_1st.Count; i++)
            {
                var item = data.Scores_1st[i];
                Assert.AreEqual(item.SubjectName, tmp[i][0]);
                Assert.AreEqual(item.Score, tmp[i][1]);
                Assert.AreEqual(item.Ranking, tmp[i][2]);
                Assert.AreEqual(item.ExamTime, tmp[i][3]);
            }

            object[][] tmp2 =
            [
                ["����", 104f, 2, DateTime.Parse("2000/6/6")],
                ["��ѧ", 105f, 5, DateTime.Parse("2000/6/7")],
                ["Ӣ��", 106f, 4, DateTime.Parse("2000/6/8")],
                ["��ѧ", 107f, 1, DateTime.Parse("2000/6/9")],
                ["����", 108f, 6, DateTime.Parse("2000/6/10")],
                ["����", 109f, 7, DateTime.Parse("2000/6/11")],
                ["����", 110f, 7, DateTime.Parse("2000/6/12")],
                ["����", 111f, 9, DateTime.Parse("2000/6/13")],
                ["��ʷ", 112f, 6, DateTime.Parse("2000/6/14")],
            ];

            for (int i = 0; i < data.Scores_2nd.Count; i++)
            {
                var item = data.Scores_2nd[i];
                Assert.AreEqual(item.SubjectName, tmp2[i][0]);
                Assert.AreEqual(item.Score, tmp2[i][1]);
                Assert.AreEqual(item.Ranking, tmp2[i][2]);
                Assert.AreEqual(item.ExamTime, tmp2[i][3]);
            }
        }
    }
}