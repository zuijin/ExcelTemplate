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

        public T Data => _data;
        internal Dictionary<string, (int row, int col)> _formFieldDic = new Dictionary<string, (int row, int col)>();
        internal Dictionary<object, int> _listDic = new Dictionary<object, int>();

        public HintBuilder(TemplateDesign design, IWorkbook workbook, T data, List<CellException> exceptions)
        {
            _exceptions = new List<CellException>();
            _design = design;
            _workbook = workbook;
            _data = data;
            _exceptions.AddRange(exceptions);

            InitDic();
        }

        private void InitDic()
        {
            throw new NotImplementedException();
        }

        public void AddError(int row, int col, string message)
        {
            _exceptions.Add(new CellException(row, col, message));
        }

        public FieldHintExp<T, TField> For<TField>(Expression<Func<T, TField>> expression)
        {
            return new FieldHintExp<T, TField>(this, expression);
        }
    }
}
