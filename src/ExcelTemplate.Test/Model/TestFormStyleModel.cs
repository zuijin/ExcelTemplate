using ExcelTemplate.Attributes;
using ExcelTemplate.Style;

namespace ExcelTemplate.Test.Model
{
    [StyleDic(Key = "date_format", DataFormat = "yyyy/m/d h:mm:ss")]
    [StyleDic(Key = "title_red", BackgroundColor = "FF0000", FontHeightInPoints = 18, IsBold = true, Alignment = HorizontalAlignment.Center)]
    internal class TestFormStyleModel
    {
        [Title("姓名：", "B2", Style = "title_red")]
        [Position("C2", Style = "title_red")]
        public string StudentName { get; set; }

        [Title("性别：", "D2")]
        [Position("E2")]
        [Style(BackgroundColor = "0000FF", FontHeightInPoints = 12, IsBold = false, Alignment = HorizontalAlignment.Left)]
        public string Sex { get; set; }

        [Title("生日：", "B3")]
        [Position("C3", Style = "date_format")]
        public DateTime BirthDate { get; set; }

    }
}
