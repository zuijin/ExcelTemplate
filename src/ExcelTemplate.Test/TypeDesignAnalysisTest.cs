using ExcelTemplate.Model;
using ExcelTemplate.Test.Model;
using NPOI.SS.Formula.Functions;

namespace ExcelTemplate.Test
{
    public class TypeDesignAnalysisTest
    {
        [Fact]
        public void StyleTest()
        {
            var design = TypeDesignAnalysis.DesignAnalysis(typeof(TestFormStyleModel));

            EachBlock<ValueBlock>(design.BlockSection, (block) =>
            {
                if (block.FieldPath == "StudentName")
                {
                    Assert.NotNull(block.Style);
                    Assert.Equal("FF0000", block.Style.FillBackgroundColor);
                    Assert.Equal(18, block.Style.Font.FontHeightInPoints);
                    Assert.True(block.Style.Font.IsBold);
                    Assert.Equal(NPOI.SS.UserModel.HorizontalAlignment.Center, block.Style.Alignment);
                }

                if (block.FieldPath == "Sex")
                {
                    Assert.NotNull(block.Style);
                    Assert.Equal("0000FF", block.Style.FillBackgroundColor);
                    Assert.Equal(12, block.Style.Font.FontHeightInPoints);
                    Assert.Equal(NPOI.SS.UserModel.HorizontalAlignment.Left, block.Style.Alignment);
                }

                if (block.FieldPath == "BirthDate")
                {
                    Assert.NotNull(block.Style);
                    Assert.Equal("yyyy/m/d h:mm:ss", block.Style.DataFormat);
                }
            });

            EachBlock<TextBlock>(design.BlockSection, (block) =>
            {
                if (block.Text == "姓名：")
                {
                    Assert.NotNull(block.Style);
                    Assert.Equal("FF0000", block.Style.FillBackgroundColor);
                    Assert.Equal(18, block.Style.Font.FontHeightInPoints);
                    Assert.True(block.Style.Font.IsBold);
                    Assert.Equal(NPOI.SS.UserModel.HorizontalAlignment.Center, block.Style.Alignment);
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
