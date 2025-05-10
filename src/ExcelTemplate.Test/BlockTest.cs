using ExcelTemplate.Model;
using ExcelTemplate.Style;
using NPOI.SS.Formula.Functions;

namespace ExcelTemplate.Test
{
    [TestClass]
    public class BlockTest
    {
        [TestMethod]
        public void TestBlockClone()
        {
            var textBlock = new TextBlock()
            {
                Position = "B2",
                MergeTo = "C3",
                Text = "aaa",
                Style = new Style.ETStyle(),
            };

            var cloneBlock = (TextBlock)textBlock.Clone();

            Assert.AreNotSame(textBlock, cloneBlock);
            Assert.AreNotSame(textBlock.Position, cloneBlock.Position);
            Assert.AreNotSame(textBlock.MergeTo, cloneBlock.MergeTo);
            Assert.AreNotSame(textBlock.Style, cloneBlock.Style);
        }

        [TestMethod]
        public void TestBlockClone2()
        {
            var textBlock = new TestBlock()
            {
                Position = "B2",
                MergeTo = "C3",
                Text = "aaa",
                testPosition = "D4"
            };

            var cloneBlock = (TestBlock)textBlock.Clone();

            Assert.AreSame(textBlock.testPosition, cloneBlock.testPosition);
        }

        [TestMethod]
        public void TestBlockClone3()
        {
            var textBlock = new TestBlock2()
            {
                Position = "B2",
                MergeTo = "C3",
                Text = "aaa",
                testPosition = "D4"
            };

            var cloneBlock = (TestBlock2)textBlock.Clone();

            Assert.AreNotSame(textBlock.testPosition, cloneBlock.testPosition);
        }

        public class TestBlock : TextBlock
        {
            public Position testPosition { get; set; }
        }

        public class TestBlock2 : TextBlock
        {
            public Position testPosition { get; set; }

            public override object Clone()
            {
                var obj = (TestBlock2)base.Clone();
                obj.testPosition = (Position)this.testPosition.Clone();

                return obj;
            }
        }
    }
}