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

            var template = Template.FromType(typeof(MixtureModel));
            var builder = template.GetHintBuilder<MixtureModel>(file);

            foreach (var item in builder.Data.Scores_1st)
            {
                builder.For(a => a.Scores_1st.Pick(item).Score).AddMessage("错误的分数");
            }

            for (var i = 0; i < builder.Data.Scores_2nd.Count; i++)
            {
                if (i % 2 == 0)
                {
                    builder.For(a => a.Scores_2nd.Pick(i).ExamTime).AddMessage("错误的时间");
                }
            }

            builder.For(a => a.TotalRanking_2nd).AddMessage("错误的排名");

            builder.SetMessageBgColor("FF0000");
            var workbook = builder.BuildExcel();
            workbook.Save("Temp/hit.xlsx");
        }

        public class PersonInfo
        {
            public string Name { get; set; }

            public List<OrderInfo> Orders { get; set; }
        }

        public class OrderInfo
        {
            public string OrderId { get; set; }

            public decimal Price { get; set; }
        }

        //[TestMethod]
        public void BuilderExpressionTest2()
        {
            var templateFileName = "...";
            var template = Template.FromExcel(templateFileName);

            var dataFileName = "...";
            var builder = template.GetHintBuilder<PersonInfo>(dataFileName);

            // 普通字段，直接添加错误提示
            builder.For(a => a.Name).AddMessage("名称错误");

            // （第一种）数组字段，foreach 通过 Pick() 方法定位
            foreach (var item in builder.Data.Orders)
            {
                builder.For(a => a.Orders.Pick(item).OrderId).AddMessage("错误的订单id");
            }

            // （第二种）数组字段，for 通过 Pick() 方法定位
            for (var i = 0; i < builder.Data.Orders.Count; i++)
            {
                builder.For(a => a.Orders.Pick(i).OrderId).AddMessage("错误的订单id");
            }

            // 错误提示颜色
            builder.SetMessageBgColor("FF0000");

            var workbook = builder.BuildExcel();
            workbook.Save("Temp/hit.xlsx");
        }
    }
}
