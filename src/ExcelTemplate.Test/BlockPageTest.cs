using ExcelTemplate.Model;
using ExcelTemplate.Test.Model;
using NPOI.SS.Formula.Functions;

namespace ExcelTemplate.Test
{
    public class BlockPageTest
    {
        [Fact]
        public void TestBlockPageClone()
        {
            var design = TypeDesignAnalysis.DesignAnalysis(typeof(FromModel));
            var clone = (BlockPage)design.BlockPage.Clone();

            var currentClone = clone;
            var currentOrigin = design.BlockPage;

            while (currentClone != null)
            {
                Assert.NotSame(currentClone, currentOrigin);
                Assert.NotSame(currentClone.RowBlocks, currentOrigin.RowBlocks);

                for (var i = 0; i < clone.RowBlocks.Count; i++)
                {
                    Assert.NotSame(clone.RowBlocks[i], currentOrigin.RowBlocks[i]);
                }

                currentClone = currentClone.Next;
                currentOrigin = currentOrigin.Next;
            }
        }
    }
}