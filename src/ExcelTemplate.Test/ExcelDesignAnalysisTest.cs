using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTemplate.Test
{
    [TestClass]
    public class ExcelDesignAnalysisTest
    {
        [TestMethod]
        public void DesignAnalysisTest()
        {
            var filePath = "Files/Mixture_Template.xlsx";
            var design = new ExcelDesignAnalysis().DesignAnalysis(filePath);
        }
    }
}
