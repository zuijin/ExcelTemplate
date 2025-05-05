using ExcelTemplate.Test.Model;

namespace ExcelTemplate.Test
{
    public class FormTest
    {
        [Fact]
        public void TestReadForm()
        {
            var filePath = "Files/Form.xlsx";
            var file = File.Open(filePath, FileMode.Open);
            var template = TemplateCapture.Create(typeof(FormModel));

            dynamic data = template.Capture<FormModel>(file);
            Assert.Equal(data.Field_1, 123);
            Assert.Equal(data.Field_2, 456);
            Assert.Equal(data.Field_3, "aabcc");
            Assert.Equal(data.Field_4, DateTime.Parse("2025/1/2"));
        }
    }
}