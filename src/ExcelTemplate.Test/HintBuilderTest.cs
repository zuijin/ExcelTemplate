using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTemplate.Test.Model;
using ExcelTemplate.Extensions;

namespace ExcelTemplate.Test
{
    [TestClass]
    public class HintBuilderTest
    {
        [TestMethod]
        public void BuilderExpressionTest()
        {
            var filePath = "Files/HitBuilderTest.xlsx";
            var file = File.Open(filePath, FileMode.Open);
            var template = TemplateCapture.Create(typeof(MixtureModel));

            var builder = template.CaptureHintBuilder<MixtureModel>(file);
            foreach (var item in builder.Data.Scores_1st)
            {
                builder.For(a => a.Scores_1st.Pick(item).Score).AddError("aaaa");
            }

            foreach (var item in builder.Data.Scores_2nd)
            {
                builder.For(a => a.Scores_2nd.Pick(item).Score).AddError("bbb");
            }

            builder.SetErrorBgColor("FF0000");
            var workbook = builder.BuildErrorExcel();
            workbook.Save("Temp/hit.xlsx");
        }
    }
}
