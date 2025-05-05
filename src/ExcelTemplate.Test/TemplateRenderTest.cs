using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelTemplate.Model;
using ExcelTemplate.Test.Model;
using ExcelTemplate.Extensions;
using System.Reflection;
using ExcelTemplate.Attributes;
using Microsoft.VisualStudio.TestPlatform.Common.DataCollection;
using ExcelTemplate.Helper;

namespace ExcelTemplate.Test
{
    public class TemplateRenderTest
    {
        /// <summary>
        /// 渲染表单
        /// </summary>
        [Fact]
        public void RenderFormTest()
        {
            var render = TemplateRender.Create(typeof(FormModel));
            var data = new FormModel()
            {
                Field_1 = 555,
                Field_2 = 999,
                Field_3 = "字符串测试",
                Field_4 = DateTime.Parse("2025-05-05 11:55:56"),
            };

            var workbook = render.Render(data);
            var sheet = workbook.GetSheetAt(0);

            var props = typeof(FormModel).GetProperties();
            foreach (var prop in props)
            {
                var titleAttrs = prop.GetCustomAttributes<TitleAttribute>();
                foreach (var attr in titleAttrs)
                {
                    var cell = sheet.GetCell(attr.Position);
                    Assert.Equal(attr.Title, cell.GetValue().ToString());
                }

                var posAttrs = prop.GetCustomAttributes<PositionAttribute>();
                foreach (var attr in posAttrs)
                {
                    var cell = sheet.GetCell(attr.Position);
                    var cellValue = Convert.ChangeType(cell.GetValue(), prop.PropertyType);
                    var objValue = prop.GetValue(data);

                    Assert.Equal(objValue, cellValue);
                }
            }
        }

        /// <summary>
        /// 渲染列表
        /// </summary>
        [Fact]
        public void RenderListTest()
        {
            var render = TemplateRender.Create(typeof(ListModel));
            var data = new ListModel()
            {
                Children = new List<ListModel.ListItem>()
                {
                    new ListModel.ListItem()
                    {
                        Field_1 = 555,
                        Field_2 = 999,
                        Field_3 = "字符串测试",
                        Field_4 = DateTime.Parse("2025-05-05 11:55:56"),
                    },
                    new ListModel.ListItem()
                    {
                        Field_1 = 555,
                        Field_2 = 999,
                        Field_3 = "字符串测试",
                        Field_4 = DateTime.Parse("2025-05-05 11:55:56"),
                    },
                    new ListModel.ListItem()
                    {
                        Field_1 = 555,
                        Field_2 = 999,
                        Field_3 = "字符串测试",
                        Field_4 = DateTime.Parse("2025-05-05 11:55:56"),
                    }
                 }
            };

            var workbook = render.Render(data);
            var sheet = workbook.GetSheetAt(0);
            var props = typeof(ListModel).GetProperties();
            var childrenProp = props.First(a => a.Name == nameof(data.Children));
            var titleRowCount = typeof(ListModel.ListItem).GetProperties().Select(a => a.GetCustomAttribute<MergeAttribute>()).Max(a => a?.Titles.Length ?? 0) + 1;
            var currentRow = new Position(childrenProp.GetCustomAttribute<PositionAttribute>().Position).Row + titleRowCount;

            foreach (var item in data.Children)
            {
                var row = sheet.GetRow(currentRow);

                foreach (var prop in item.GetType().GetProperties())
                {
                    var titleAttr = prop.GetCustomAttribute<TitleAttribute>();
                    var col = new Position(titleAttr.Position).Col;
                    var cell = sheet.GetCell(currentRow, col);
                    var cellValue = Convert.ChangeType(cell.GetValue(), prop.PropertyType);
                    var objValue = prop.GetValue(item);

                    Assert.Equal(objValue, cellValue);
                }

                currentRow++;
            }
        }

        /// <summary>
        /// 渲染合并表头的列表
        /// </summary>
        [Fact]
        public void RenderMergeHeaderListTest()
        {
            var render = TemplateRender.Create(typeof(MergeHeaderListModel));
            var data = new MergeHeaderListModel()
            {
                Children = new List<MergeHeaderListModel.ListItem>()
                {
                    new MergeHeaderListModel.ListItem()
                    {
                        Field_1 = 555,
                        Field_2 = 999,
                        Field_3 = "字符串测试",
                        Field_4 = DateTime.Parse("2025-05-05 11:55:56"),
                    },
                    new MergeHeaderListModel.ListItem()
                    {
                        Field_1 = 555,
                        Field_2 = 999,
                        Field_3 = "字符串测试",
                        Field_4 = DateTime.Parse("2025-05-05 11:55:56"),
                    },
                    new MergeHeaderListModel.ListItem()
                    {
                        Field_1 = 555,
                        Field_2 = 999,
                        Field_3 = "字符串测试",
                        Field_4 = DateTime.Parse("2025-05-05 11:55:56"),
                    }
                 }
            };

            var workbook = render.Render(data);
            //workbook.Save("Temp/mergeheaderlist.xlsx");

            var sheet = workbook.GetSheetAt(0);
            var props = typeof(MergeHeaderListModel).GetProperties();
            var childrenProp = props.First(a => a.Name == nameof(data.Children));
            var titleRowCount = typeof(MergeHeaderListModel.ListItem).GetProperties().Select(a => a.GetCustomAttribute<MergeAttribute>()).Max(a => a?.Titles.Length ?? 0) + 1;
            var currentRow = new Position(childrenProp.GetCustomAttribute<PositionAttribute>().Position).Row + titleRowCount;

            foreach (var item in data.Children)
            {
                var row = sheet.GetRow(currentRow);

                foreach (var prop in item.GetType().GetProperties())
                {
                    var titleAttr = prop.GetCustomAttribute<TitleAttribute>();
                    var col = new Position(titleAttr.Position).Col;
                    var cell = sheet.GetCell(currentRow, col);
                    var cellValue = Convert.ChangeType(cell.GetValue(), prop.PropertyType);
                    var objValue = prop.GetValue(item);

                    Assert.Equal(objValue, cellValue);
                }

                currentRow++;
            }
        }
    }
}
