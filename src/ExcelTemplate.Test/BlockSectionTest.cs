using ExcelTemplate.Model;
using ExcelTemplate.Test.Model;
using NPOI.SS.Formula.Functions;

namespace ExcelTemplate.Test
{
    public class BlockSectionTest
    {
        [Fact]
        public void TestBlockSectionClone()
        {
            var design = TypeDesignAnalysis.DesignAnalysis(typeof(FromModel));
            var clone = (BlockSection)design.BlockSection.Clone();

            var currentClone = clone;
            var currentOrigin = design.BlockSection;

            while (currentClone != null)
            {
                Assert.NotSame(currentClone, currentOrigin);
                Assert.NotSame(currentClone.Blocks, currentOrigin.Blocks);

                for (var i = 0; i < clone.Blocks.Count; i++)
                {
                    Assert.NotSame(clone.Blocks[i], currentOrigin.Blocks[i]);
                }

                currentClone = currentClone.Next;
                currentOrigin = currentOrigin.Next;
            }
        }
    }
}