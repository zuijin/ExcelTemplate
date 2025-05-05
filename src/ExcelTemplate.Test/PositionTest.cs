using ExcelTemplate.Model;

namespace ExcelTemplate.Test
{
    public class PositionTest
    {
        [Fact]
        public void NewPositionTest()
        {
            try
            {
                Position p = new Position("12");
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.Equal("letter格式错误", ex.Message);
            }
        }

        [Fact]
        public void NewPositionTest2()
        {
            try
            {
                Position p = new Position("A");
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.Equal("letter格式错误", ex.Message);
            }
        }

        [Fact]
        public void NewPositionTest3()
        {
            try
            {
                Position p = new Position(-1, 0);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.Equal("row不能小于0", ex.Message);
            }
        }

        [Fact]
        public void NewPositionTest4()
        {
            try
            {
                Position p = new Position(0, -1);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.Equal("col不能小于0", ex.Message);
            }
        }

        [Fact]
        public void ChangeFieldTest()
        {
            try
            {
                Position p = new Position("A1");
                p.Letter = "B";
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.Equal("letter格式错误", ex.Message);
            }
        }

        [Fact]
        public void ChangeFieldTest2()
        {
            try
            {
                Position p = new Position("A1");
                p.Row = -1;
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.Equal("row不能小于0", ex.Message);
            }
        }

        [Fact]
        public void ChangeFieldTest3()
        {
            try
            {
                Position p = new Position("A1");
                p.Col = -1;
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.Equal("col不能小于0", ex.Message);
            }
        }


        [Fact]
        public void ValueTest()
        {
            Position p = new Position("A1");
            Assert.Equal(0, p.Row);
            Assert.Equal(0, p.Col);

            Position p2 = new Position(2, 2);
            Assert.Equal("C3", p2.Letter);

            Position p3 = p2.GetOffset(1, 0);
            Assert.Equal("C4", p3.Letter);

            Position p4 = p3.GetOffset(0, 1);
            Assert.Equal("D4", p4.Letter);

            try
            {
                Position p5 = p.GetOffset(0, -1);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.Equal("col不能小于0", ex.Message);
            }

            try
            {
                Position p6 = p.GetOffset(-1, 0);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.Equal("row不能小于0", ex.Message);
            }
        }


        [Fact]
        public void ValueTest2()
        {
            Position p = new Position("AA1");
            Assert.Equal(26 + 1 - 1, p.Col);

            Position p2 = new Position("ABC255");
            Assert.Equal(254, p2.Row);
            Assert.Equal((1 * 26 * 26) + (2 * 26) + 3 - 1, p2.Col);

            Position p3 = new Position("BCD1");
            Assert.Equal((2 * 26 * 26) + (3 * 26) + 4 - 1, p3.Col);

            Position p4 = new Position(0, (2 * 26 * 26) + (3 * 26) + 4 - 1);
            Assert.Equal("BCD1", p4.Letter);
        }


        /// <summary>
        /// 大小写测试
        /// </summary>
        [Fact]
        public void IgnoreCaseTest()
        {
            Position p = new Position("a1");
            Assert.Equal(0, p.Row);
            Assert.Equal(0, p.Col);

            p.Letter = "e5";
            Assert.Equal(4, p.Row);
            Assert.Equal(4, p.Col);

        }
    }
}
