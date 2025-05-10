using ExcelTemplate.Model;

namespace ExcelTemplate.Test
{
    [TestClass]
    public class PositionTest
    {
        [TestMethod]
        public void NewPositionTest()
        {
            try
            {
                Position p = new Position("12");
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.AreEqual("letter格式错误", ex.Message);
            }
        }

        [TestMethod]
        public void NewPositionTest2()
        {
            try
            {
                Position p = new Position("A");
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.AreEqual("letter格式错误", ex.Message);
            }
        }

        [TestMethod]
        public void NewPositionTest3()
        {
            try
            {
                Position p = new Position(-1, 0);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.AreEqual("row不能小于0", ex.Message);
            }
        }

        [TestMethod]
        public void NewPositionTest4()
        {
            try
            {
                Position p = new Position(0, -1);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.AreEqual("col不能小于0", ex.Message);
            }
        }

        [TestMethod]
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
                Assert.AreEqual("letter格式错误", ex.Message);
            }
        }

        [TestMethod]
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
                Assert.AreEqual("row不能小于0", ex.Message);
            }
        }

        [TestMethod]
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
                Assert.AreEqual("col不能小于0", ex.Message);
            }
        }


        [TestMethod]
        public void ValueTest()
        {
            Position p = new Position("A1");
            Assert.AreEqual(0, p.Row);
            Assert.AreEqual(0, p.Col);

            Position p2 = new Position(2, 2);
            Assert.AreEqual("C3", p2.Letter);

            Position p3 = p2.GetOffset(1, 0);
            Assert.AreEqual("C4", p3.Letter);

            Position p4 = p3.GetOffset(0, 1);
            Assert.AreEqual("D4", p4.Letter);

            try
            {
                Position p5 = p.GetOffset(0, -1);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.AreEqual("col不能小于0", ex.Message);
            }

            try
            {
                Position p6 = p.GetOffset(-1, 0);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.AreEqual("row不能小于0", ex.Message);
            }
        }


        [TestMethod]
        public void ValueTest2()
        {
            Position p = new Position("AA1");
            Assert.AreEqual(26 + 1 - 1, p.Col);

            Position p2 = new Position("ABC255");
            Assert.AreEqual(254, p2.Row);
            Assert.AreEqual((1 * 26 * 26) + (2 * 26) + 3 - 1, p2.Col);

            Position p3 = new Position("BCD1");
            Assert.AreEqual((2 * 26 * 26) + (3 * 26) + 4 - 1, p3.Col);

            Position p4 = new Position(0, (2 * 26 * 26) + (3 * 26) + 4 - 1);
            Assert.AreEqual("BCD1", p4.Letter);
        }


        /// <summary>
        /// 大小写测试
        /// </summary>
        [TestMethod]
        public void IgnoreCaseTest()
        {
            Position p = new Position("a1");
            Assert.AreEqual(0, p.Row);
            Assert.AreEqual(0, p.Col);

            p.Letter = "e5";
            Assert.AreEqual(4, p.Row);
            Assert.AreEqual(4, p.Col);

        }
    }
}
