using ExcelTemplate.Hint;
using ExcelTemplate.Model;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

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

        public Template(TemplateDesign design)
        {
            _design = design;
            _capture = new TemplateCapture(_design);
            _render = new TemplateRender(_design);
        }

        public static Template FromType<T>()
        {
            var design = new TypeDesignAnalysis().DesignAnalysis(typeof(T));
            return new Template(design);
        }

        public static Template FromType(Type type)
        {
            var design = new TypeDesignAnalysis().DesignAnalysis(type);
            return new Template(design);
        }

        public static Template FromExcel(string excelFile)
        {
            var design = new ExcelDesignAnalysis().DesignAnalysis(excelFile);
            return new Template(design);
        }

        /// <summary>
        /// 从Excel读取数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream"></param>
        /// <returns></returns>
        public T Capture<T>(Stream stream)
        {
            return _capture.Capture<T>(stream);
        }

        /// <summary>
        /// 从Excel读取数据
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="dataType">数据类型</param>
        /// <returns></returns>
        public object Capture(Stream stream, Type dataType)
        {
            return _capture.Capture(stream, dataType);
        }

        /// <summary>
        /// 获取提示信息生成器
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream"></param>
        /// <returns></returns>
        public HintBuilder<T> GetHintBuilder<T>(Stream stream)
        {
            return _capture.GetHintBuilder<T>(stream);
        }

        public HintBuilder<T> GetHintBuilder<T>(string fileName)
        {
            using var stream = File.OpenRead(fileName);
            return _capture.GetHintBuilder<T>(stream);
        }

        /// <summary>
        /// 将数据写入Excel
        /// </summary>
        /// <param name="data"></param>
        /// <param name="workbook"></param>
        public void Render(object data, IWorkbook workbook)
        {
            _render.Render(data, workbook);
        }

        /// <summary>
        /// 将数据写入Excel
        /// </summary>
        /// <param name="data"></param>
        /// <param name="stream">excel文件流</param>
        public void Render(object data, Stream stream)
        {
            _render.Render(data, stream);
        }

        /// <summary>
        /// 从数据生成一个Excel
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public IWorkbook Render(object data)
        {
            return _render.Render(data);
        }

        /// <summary>
        /// 添加读取Excel时的值映射方法
        /// </summary>
        /// <param name="fieldPath"></param>
        /// <param name="mappingFunc"></param>
        /// <exception cref="Exception"></exception>
        public void AddReadMapping(string fieldPath, Func<object, object> mappingFunc)
        {
            _capture.AddMapping(fieldPath, mappingFunc);
        }

        /// <summary>
        /// 添加写入Excel时的值映射方法
        /// </summary>
        /// <param name="fieldPath"></param>
        /// <param name="mappingFunc"></param>
        public void AddRenderMapping(string fieldPath, Func<object, object> mappingFunc)
        {
            _render.AddMapping(fieldPath, mappingFunc);
        }
    }
}
