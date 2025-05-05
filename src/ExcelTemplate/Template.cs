using System;
using ExcelTemplate.Model;
using NPOI.SS.UserModel;

namespace ExcelTemplate
{
    /// <summary>
    /// Excel 模版对象
    /// </summary>
    public class Template
    {
        TemplateDesign _design;
        TemplateCapture _capture;
        TemplateRender _render;

        public Template(Type designType)
        {
            _design = TypeDesignAnalysis.DesignAnalysis(designType);
            _capture = new TemplateCapture(_design);
            _render = new TemplateRender(_design);
        }

        /// <summary>
        /// 从Excel读取数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public T ReadFromExcel<T>(IWorkbook workbook)
        {
            return _capture.Capture<T>(workbook);
        }

        /// <summary>
        /// 从Excel读取数据
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="dataType">数据类型</param>
        /// <returns></returns>
        public object ReadFromExcel(IWorkbook workbook, Type dataType)
        {
            return _capture.Capture(workbook, dataType);
        }

        /// <summary>
        /// 将数据写入Excel
        /// </summary>
        /// <param name="data"></param>
        /// <param name="workbook"></param>
        public void WriteToExcel(object data, IWorkbook workbook)
        {
            _render.Render(data, workbook);
        }

        /// <summary>
        /// 从数据生成一个Excel
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public IWorkbook GetExcelFromData(object data)
        {
            return _render.Render(data);
        }
    }
}
