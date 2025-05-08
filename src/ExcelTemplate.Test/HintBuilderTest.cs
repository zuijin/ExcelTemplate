using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTemplate.Test.Model;
using ExcelTemplate.Hint;

namespace ExcelTemplate.Test
{
    public class HintBuilderTest
    {
        //[Fact]
        public void BuilderExpressionTest()
        {
            var filePath = "Files/Mixture.xlsx";
            var file = File.Open(filePath, FileMode.Open);
            var template = TemplateCapture.Create(typeof(MixtureModel));

            var builder = template.CaptureHintBuilder<MixtureModel>(file);
            var i = 0;
            foreach (var item in builder.Data.Scores_1st)
            {
                i++;
                builder.For(a => a.Scores_1st.Pick(i).Score).AddError("aaaa");
                builder.For(a => a.Sex).AddError("bbb");
            }

        }
    }
}
