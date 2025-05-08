using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Text;
using ExcelTemplate.Model;
using NPOI.SS.UserModel;
using NPOI.Util;

namespace ExcelTemplate.Hint
{
    public class HintBuilder<T>
    {
        private T _data;
        private List<CellException> _exceptions;
        private IWorkbook _workbook;
        private TemplateDesign _design;

        internal Dictionary<string, (int row, int col)> _formFieldDic = new Dictionary<string, (int row, int col)>();
        internal Dictionary<object, Dictionary<string, (int row, int col)>> _listDic = new Dictionary<object, Dictionary<string, (int row, int col)>>();

        public HintBuilder(TemplateDesign design, IWorkbook workbook, T data, List<CellException> exceptions)
        {
            _exceptions = new List<CellException>();
            _design = design;
            _workbook = workbook;
            _data = data;
            _exceptions.AddRange(exceptions);
        }

        private void InitDic()
        {
            new HintBuilder<TemplateDesign>(new TemplateDesign(), _workbook, new TemplateDesign(), null).For(a => a.SourceType).AddError("");
        }

        public void AddError(int row, int col, string message)
        {
            _exceptions.Add(new CellException(row, col, message));
        }

        public FieldHintExecutor<T, TField> For<TField>(Expression<Func<T, TField>> expression)
        {
            return new FieldHintExecutor<T,TField>(this, expression);
        }

        public FieldHintExecutor<T, IEnumerable<TField>> For<TField>(Expression<Func<T, IEnumerable<TField>>> expression)
        {
            return new FieldHintExecutor<T, IEnumerable<TField>>(this, expression);
        }
    }
}
