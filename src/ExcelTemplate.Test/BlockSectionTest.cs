using ExcelTemplate.Model;
using ExcelTemplate.Test.Model;
using NPOI.SS.Formula.Functions;

namespace ExcelTemplate.Test
{
    [TestClass]
    public class BlockSectionTest
    {
        [TestMethod]
        public void TestBlockSectionClone()
        {
            var design = TypeDesignAnalysis.DesignAnalysis(typeof(FormModel));
            var clone = (BlockSection)design.BlockSection.Clone();

            var currentClone = clone;
            var currentOrigin = design.BlockSection;

            while (currentClone != null)
            {
                Assert.AreNotSame(currentClone, currentOrigin);
                Assert.AreNotSame(currentClone.Blocks, currentOrigin.Blocks);

                for (var i = 0; i < clone.Blocks.Count; i++)
                {
                    Assert.AreNotSame(clone.Blocks[i], currentOrigin.Blocks[i]);
                }

                currentClone = currentClone.Next;
                currentOrigin = currentOrigin.Next;
            }
        }
    }
}