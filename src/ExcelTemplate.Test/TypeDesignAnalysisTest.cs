using ExcelTemplate.Model;
using ExcelTemplate.Test.Model;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace ExcelTemplate.Test
{
    [TestClass]
    public class TypeDesignAnalysisTest
    {
        [TestMethod]
        public void StyleTest()
        {
            var design = new TypeDesignAnalysis().DesignAnalysis(typeof(TestFormStyleModel));

            EachBlock<ValueBlock>(design.BlockSection, (block) =>
            {
                if (block.FieldPath == "StudentName")
                {
                    Assert.IsNotNull(block.Style);
                    Assert.AreEqual("FF0000", block.Style.Font.Color);
                    Assert.AreEqual(18, block.Style.Font.FontHeightInPoints);
                    Assert.IsTrue(block.Style.Font.IsBold);
                    Assert.AreEqual(NPOI.SS.UserModel.HorizontalAlignment.Center, block.Style.Alignment);
                }

                if (block.FieldPath == "Sex")
                {
                    Assert.IsNotNull(block.Style);
                    Assert.AreEqual("0000FF", block.Style.FillBackgroundColor);
                    Assert.AreEqual("0000FF", block.Style.FillForegroundColor);
                    Assert.AreEqual(FillPattern.SolidForeground, block.Style.FillPattern);
                    Assert.AreEqual(12, block.Style.Font.FontHeightInPoints);
                    Assert.AreEqual(NPOI.SS.UserModel.HorizontalAlignment.Left, block.Style.Alignment);
                }

                if (block.FieldPath == "BirthDate")
                {
                    Assert.IsNotNull(block.Style);
                    Assert.AreEqual("yyyy/m/d h:mm:ss", block.Style.DataFormat);
                }
            });

            EachBlock<TextBlock>(design.BlockSection, (block) =>
            {
                if (block.Text == "姓名：")
                {
                    Assert.IsNotNull(block.Style);
                    Assert.AreEqual("FF0000", block.Style.Font.Color);
                    Assert.AreEqual(18, block.Style.Font.FontHeightInPoints);
                    Assert.IsTrue(block.Style.Font.IsBold);
                    Assert.AreEqual(NPOI.SS.UserModel.HorizontalAlignment.Center, block.Style.Alignment);
                }
            });
        }


        private void EachBlock<T>(BlockSection section, Action<T> func)
        {
            while (section != null)
            {
                foreach (var item in section.Blocks.OfType<T>())
                {
                    func?.Invoke(item);
                }

                section = section.Next;
            }
        }
    }
}
