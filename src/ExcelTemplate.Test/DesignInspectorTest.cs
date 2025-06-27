using ExcelTemplate.Attributes;
using ExcelTemplate.Exceptions;

namespace ExcelTemplate.Test;

[TestClass]
public class DesignInspectorTest
{
    [TestMethod]
    public void CheckTest()
    {
        var design = new TypeDesignAnalysis().DesignAnalysis(typeof(Obj1));

        try
        {
            DesignInspector.Check(design);
            Assert.Fail();
        }
        catch (TemplateDesignException ex)
        {
            Assert.AreEqual(TemplateDesignExceptionType.PositionConflict, ex.ErrorType);
        }
    }

    public class Obj1
    {
        [Position("C1")]
        public int Field_1 { get; set; }

        [Position("C1")]
        public int Field_2 { get; set; }


    }
}
