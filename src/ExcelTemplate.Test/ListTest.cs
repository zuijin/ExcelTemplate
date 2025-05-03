using ExcelTemplate.Test.Model;

namespace ExcelTemplate.Test
{
    public class ListTest
    {
        /// <summary>
        /// 读取列表
        /// </summary>
        [Fact]
        public void TestReadList()
        {
            var filePath = "Files/List.xlsx";
            var file = File.Open(filePath, FileMode.Open);
            var template = TemplateReader.FromFile(file, typeof(ListModel));

            dynamic data = template.GetData();
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
            var template = TemplateReader.FromFile(file, typeof(MergeHeaderListModel));

            dynamic data = template.GetData();
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
    }
}