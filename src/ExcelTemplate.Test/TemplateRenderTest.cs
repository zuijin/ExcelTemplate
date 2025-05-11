using System.Reflection;
using ExcelTemplate.Attributes;
using ExcelTemplate.Extensions;
using ExcelTemplate.Style;
using ExcelTemplate.Test.Model;
using NPOI.SS.UserModel;

namespace ExcelTemplate.Test
{
    [TestClass]
    public class TemplateRenderTest
    {
        /// <summary>
        /// 渲染表单
        /// </summary>
        [TestMethod]
        public void RenderFormTest()
        {
            var render = TemplateRender.Create(typeof(FormModel));
            var data = new FormModel()
            {
                Field_1 = 555,
                Field_2 = 999,
                Field_3 = "字符串测试",
                Field_4 = DateTime.Parse("2025-05-05 11:55:56"),
            };

            var workbook = render.Render(data);
            var sheet = workbook.GetSheetAt(0);

            var props = typeof(FormModel).GetProperties();
            foreach (var prop in props)
            {
                var titleAttrs = prop.GetCustomAttributes<TitleAttribute>();
                foreach (var attr in titleAttrs)
                {
                    var cell = sheet.GetCell(attr.Position);
                    Assert.AreEqual(attr.Title, cell.GetValue().ToString());
                }

                var posAttrs = prop.GetCustomAttributes<PositionAttribute>();
                foreach (var attr in posAttrs)
                {
                    var cell = sheet.GetCell(attr.Position);
                    var cellValue = cell.GetValue(prop.PropertyType);
                    var objValue = prop.GetValue(data);

                    Assert.AreEqual(objValue, cellValue);
                }
            }
        }

        /// <summary>
        /// 渲染列表
        /// </summary>
        [TestMethod]
        public void RenderListTest()
        {
            var render = TemplateRender.Create(typeof(ListModel));
            var data = new ListModel()
            {
                Children = new List<ListModel.ListItem>()
                {
                    new ListModel.ListItem()
                    {
                        Field_1 = 555,
                        Field_2 = 999,
                        Field_3 = "字符串测试",
                        Field_4 = DateTime.Parse("2025-05-05 11:55:56"),
                    },
                    new ListModel.ListItem()
                    {
                        Field_1 = 555,
                        Field_2 = 999,
                        Field_3 = "字符串测试",
                        Field_4 = DateTime.Parse("2025-05-05 11:55:56"),
                    },
                    new ListModel.ListItem()
                    {
                        Field_1 = 555,
                        Field_2 = 999,
                        Field_3 = "字符串测试",
                        Field_4 = DateTime.Parse("2025-05-05 11:55:56"),
                    }
                 }
            };

            var workbook = render.Render(data);
            var sheet = workbook.GetSheetAt(0);
            var props = typeof(ListModel).GetProperties();
            var childrenProp = props.First(a => a.Name == nameof(data.Children));
            var titleRowCount = typeof(ListModel.ListItem).GetProperties().Select(a => a.GetCustomAttribute<MergeAttribute>()).Max(a => a?.Titles.Length ?? 0) + 1;
            var positionAttr = childrenProp.GetCustomAttribute<PositionAttribute>();
            var currentRow = positionAttr.Position.Row + titleRowCount;

            foreach (var item in data.Children)
            {
                var row = sheet.GetRow(currentRow);

                foreach (var prop in item.GetType().GetProperties())
                {
                    var colAttr = prop.GetCustomAttribute<ColAttribute>();
                    if (colAttr != null)
                    {
                        var col = positionAttr.Position.Col + colAttr.ColIndex;
                        var cell = sheet.GetCell(currentRow, col);
                        var cellValue = cell.GetValue(prop.PropertyType);
                        var objValue = prop.GetValue(item);

                        Assert.AreEqual(objValue, cellValue);
                    }
                }

                currentRow++;
            }
        }

        /// <summary>
        /// 渲染合并表头的列表
        /// </summary>
        [TestMethod]
        public void RenderMergeHeaderListTest()
        {
            var render = TemplateRender.Create(typeof(MergeHeaderListModel));
            var data = new MergeHeaderListModel()
            {
                Children = new List<MergeHeaderListModel.ListItem>()
                {
                    new MergeHeaderListModel.ListItem()
                    {
                        Field_1 = 555,
                        Field_2 = 999,
                        Field_3 = "字符串测试",
                        Field_4 = DateTime.Parse("2025-05-05 11:55:56"),
                    },
                    new MergeHeaderListModel.ListItem()
                    {
                        Field_1 = 555,
                        Field_2 = 999,
                        Field_3 = "字符串测试",
                        Field_4 = DateTime.Parse("2025-05-05 11:55:56"),
                    },
                    new MergeHeaderListModel.ListItem()
                    {
                        Field_1 = 555,
                        Field_2 = 999,
                        Field_3 = "字符串测试",
                        Field_4 = DateTime.Parse("2025-05-05 11:55:56"),
                    }
                 }
            };

            var workbook = render.Render(data);
            //workbook.Save("Temp/mergeheaderlist.xlsx");

            var sheet = workbook.GetSheetAt(0);
            var props = typeof(MergeHeaderListModel).GetProperties();
            var childrenProp = props.First(a => a.Name == nameof(data.Children));
            var titleRowCount = typeof(MergeHeaderListModel.ListItem).GetProperties().Select(a => a.GetCustomAttribute<MergeAttribute>()).Max(a => a?.Titles.Length ?? 0) + 1;
            var positionAttr = childrenProp.GetCustomAttribute<PositionAttribute>();
            var currentRow = positionAttr.Position.Row + titleRowCount;

            foreach (var item in data.Children)
            {
                var row = sheet.GetRow(currentRow);

                foreach (var prop in item.GetType().GetProperties())
                {
                    var colAttr = prop.GetCustomAttribute<ColAttribute>();
                    if (colAttr != null)
                    {
                        var col = positionAttr.Position.Col + colAttr.ColIndex;
                        var cell = sheet.GetCell(currentRow, col);
                        var cellValue = cell.GetValue(prop.PropertyType);
                        var objValue = prop.GetValue(item);

                        Assert.AreEqual(objValue, cellValue);
                    }
                }

                currentRow++;
            }
        }

        /// <summary>
        /// 渲染混排
        /// </summary>
        [TestMethod]
        public void RenderMixtureTest()
        {
            var filePath = "Files/MixtureRender.xlsx";
            var file = File.Open(filePath, FileMode.Open);
            var originWorkbook = WorkbookFactory.Create(file);

            var capture = TemplateCapture.Create(typeof(MixtureModel));
            var originData = capture.Capture<MixtureModel>(originWorkbook);

            var render = TemplateRender.Create(typeof(MixtureModel));
            var workbook = render.Render(originData);

            var originSheet = originWorkbook.GetSheetAt(0);
            var sheet = workbook.GetSheetAt(0);

            //workbook.Save("Temp/Mixture.xlsx");

            //对比单元格
            foreach (var row in sheet.AsEnumerable())
            {
                foreach (var cell in row.AsEnumerable())
                {
                    var tmpCell = originSheet.GetCell(cell.RowIndex, cell.ColumnIndex);
                    var val1 = tmpCell.GetValue();
                    var val2 = cell.GetValue();

                    //if (val1 is DateTime || val2 is DateTime)
                    //{
                    //    val1 = tmpCell.DateCellValue;
                    //    val2 = cell.DateCellValue;
                    //}

                    Assert.AreEqual(val1, val2);
                }
            }

            foreach (var row in originSheet.AsEnumerable())
            {
                foreach (var cell in row.AsEnumerable())
                {
                    var tmpCell = sheet.GetCell(cell.RowIndex, cell.ColumnIndex);
                    var val1 = tmpCell.GetValue();
                    var val2 = cell.GetValue();
                    Assert.AreEqual(val1, val2);
                }
            }

            // 对比数据
            var data = capture.Capture<MixtureModel>(workbook);

            Assert.AreEqual(originData.StudentName, data.StudentName);
            Assert.AreEqual(originData.Sex, data.Sex);
            Assert.AreEqual(originData.BirthDate, data.BirthDate);
            Assert.AreEqual(originData.TotalScore_1st, data.TotalScore_1st);
            Assert.AreEqual(originData.TotalRanking_1st, data.TotalRanking_1st);
            Assert.AreEqual(originData.TotalScore_2nd, data.TotalScore_2nd);
            Assert.AreEqual(originData.TotalRanking_2nd, data.TotalRanking_2nd);
            Assert.AreEqual(originData.Scores_1st.Count, data.Scores_1st.Count);
            Assert.AreEqual(originData.Scores_2nd.Count, data.Scores_2nd.Count);


            for (int i = 0; i < data.Scores_1st.Count; i++)
            {
                var item = data.Scores_1st[i];
                var item2 = originData.Scores_1st[i];
                Assert.AreEqual(item.SubjectName, item2.SubjectName);
                Assert.AreEqual(item.Score, item2.Score);
                Assert.AreEqual(item.Ranking, item2.Ranking);
                Assert.AreEqual(item.ExamTime, item2.ExamTime);
            }

            for (int i = 0; i < data.Scores_2nd.Count; i++)
            {
                var item = data.Scores_2nd[i];
                var item2 = originData.Scores_2nd[i];
                Assert.AreEqual(item.SubjectName, item2.SubjectName);
                Assert.AreEqual(item.Score, item2.Score);
                Assert.AreEqual(item.Ranking, item2.Ranking);
                Assert.AreEqual(item.ExamTime, item2.ExamTime);
            }
        }
        [TestMethod]
        public void RenderFormStyleTest()
        {
            var render = TemplateRender.Create(typeof(TestFormStyleModel));
            var data = new TestFormStyleModel()
            {
                StudentName = "李四",
                BirthDate = DateTime.Today,
                Sex = "女"
            };

            var builder = render.RenderHintBuilder(data);
            //builder.Workbook.Save("Temp/styletest.xlsx");
            var sheet = builder.Workbook.GetSheetAt(0);
            var position = builder.For(a => a.StudentName).GetPosition();
            var cell = sheet.GetCell(position);
            var style = ETStyleUtil.ConvertStyle(builder.Workbook, cell.CellStyle);

            Assert.AreEqual("FFFF0000", style.Font.Color);
            Assert.AreEqual(18, style.Font.FontHeightInPoints);
            Assert.AreEqual(true, style.Font.IsBold);
            Assert.AreEqual(HorizontalAlignment.Center, style.Alignment);
            Assert.AreEqual(true, style.ShrinkToFit);
            Assert.AreEqual(VerticalAlignment.Center, style.VerticalAlignment);
            Assert.AreEqual(true, style.WrapText);

            position = builder.For(a => a.BirthDate).GetPosition();
            cell = sheet.GetCell(position);
            style = ETStyleUtil.ConvertStyle(builder.Workbook, cell.CellStyle);
            Assert.AreEqual("yyyy/m/d h:mm:ss", style.DataFormat);

            position = builder.For(a => a.Sex).GetPosition();
            cell = sheet.GetCell(position);
            style = ETStyleUtil.ConvertStyle(builder.Workbook, cell.CellStyle);

            Assert.AreEqual("FF0000FF", style.FillForegroundColor);
            Assert.AreEqual(12, style.Font.FontHeightInPoints);
            Assert.AreEqual(false, style.Font.IsBold);
            Assert.AreEqual(HorizontalAlignment.Left, style.Alignment);

        }

        [TestMethod]
        public void RenderMixtureStyleTest()
        {
            var render = TemplateRender.Create(typeof(TestMixtureStyleModel));
            var date = DateTime.Parse("2024-02-03");
            var model = new TestMixtureStyleModel()
            {
                StudentName = "李四",
                BirthDate = DateTime.Parse("2025-02-03"),
                Sex = "未知",
                Scores_1st = new List<TestMixtureStyleModel.ScoreItem>()
                {
                    new TestMixtureStyleModel.ScoreItem() { SubjectName = "语文", Score = 100.5f, Ranking = 2, ExamTime = date.AddDays(1) },
                    new TestMixtureStyleModel.ScoreItem() { SubjectName = "数学", Score = 102f,   Ranking = 3, ExamTime = date.AddDays(1) },
                    new TestMixtureStyleModel.ScoreItem() { SubjectName = "英语", Score = 103,    Ranking = 4, ExamTime = date.AddDays(1) },
                    new TestMixtureStyleModel.ScoreItem() { SubjectName = "化学", Score = 104,    Ranking = 5, ExamTime = date.AddDays(1) },
                    new TestMixtureStyleModel.ScoreItem() { SubjectName = "物理", Score = 105,    Ranking = 6, ExamTime = date.AddDays(1) },
                    new TestMixtureStyleModel.ScoreItem() { SubjectName = "生物", Score = 106,    Ranking = 7, ExamTime = date.AddDays(1) },
                    new TestMixtureStyleModel.ScoreItem() { SubjectName = "地理", Score = 107,    Ranking = 8, ExamTime = date.AddDays(1) },
                    new TestMixtureStyleModel.ScoreItem() { SubjectName = "政治", Score = 108,    Ranking = 9, ExamTime = date.AddDays(1) },
                    new TestMixtureStyleModel.ScoreItem() { SubjectName = "历史", Score = 110,    Ranking = 10,ExamTime = date.AddDays(1) },
                },
                Scores_2nd = new List<TestMixtureStyleModel.ScoreItem>()
                {
                    new TestMixtureStyleModel.ScoreItem() { SubjectName = "语文", Score = 106.5f, Ranking = 1, ExamTime = date.AddDays(1) },
                    new TestMixtureStyleModel.ScoreItem() { SubjectName = "数学", Score = 107f,   Ranking = 3, ExamTime = date.AddDays(1) },
                    new TestMixtureStyleModel.ScoreItem() { SubjectName = "英语", Score = 108,    Ranking = 7, ExamTime = date.AddDays(1) },
                    new TestMixtureStyleModel.ScoreItem() { SubjectName = "化学", Score = 109,    Ranking = 5, ExamTime = date.AddDays(1) },
                    new TestMixtureStyleModel.ScoreItem() { SubjectName = "物理", Score = 110,    Ranking = 2, ExamTime = date.AddDays(1) },
                    new TestMixtureStyleModel.ScoreItem() { SubjectName = "生物", Score = 111,    Ranking = 9, ExamTime = date.AddDays(1) },
                    new TestMixtureStyleModel.ScoreItem() { SubjectName = "地理", Score = 112,    Ranking = 8, ExamTime = date.AddDays(1) },
                    new TestMixtureStyleModel.ScoreItem() { SubjectName = "政治", Score = 113,    Ranking = 11,ExamTime = date.AddDays(1) },
                    new TestMixtureStyleModel.ScoreItem() { SubjectName = "历史", Score = 114,    Ranking = 10,ExamTime = date.AddDays(1) },
                },
                TotalRanking_1st = 5,
                TotalRanking_2nd = 3,
                TotalScore_1st = 920,
                TotalScore_2nd = 968
            };

            render.AddMapping("Scores_1st.Ranking", obj => ((int)obj) + 20);

            var builder = render.RenderHintBuilder(model);
            builder.Workbook.Save("Temp/TestMixtureStyle.xlsx");
        }

    }
}
