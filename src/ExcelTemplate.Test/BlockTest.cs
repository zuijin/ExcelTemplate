using ExcelTemplate.Model;
using ExcelTemplate.Test.Model;
using NPOI.SS.Formula.Functions;

namespace ExcelTemplate.Test
{
    public class BlockTest
    {
        [Fact]
        public void TestBlockClone()
        {
            var textBlock = new TextBlock()
            {
                Position = "B2",
                MergeTo = "C3",
                Text = "aaa"
            };

            var cloneBlock = (TextBlock)textBlock.Clone();

            Assert.NotSame(textBlock, cloneBlock);
            Assert.NotSame(textBlock.Position, cloneBlock.Position);
            Assert.NotSame(textBlock.MergeTo, cloneBlock.MergeTo);
        }

        [Fact]
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

            Assert.Same(textBlock.testPosition, cloneBlock.testPosition);
        }

        [Fact]
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

            Assert.NotSame(textBlock.testPosition, cloneBlock.testPosition);
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