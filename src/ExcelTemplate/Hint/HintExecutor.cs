using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using ExcelTemplate.Model;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace ExcelTemplate.Hint
{
    public class HintExecutor
    {

    }

    public class FieldHintExecutor<T, T2>
    {
        private HintBuilder<T> _builder;
        private Expression _expression;

        public FieldHintExecutor(HintBuilder<T> builder, Expression<Func<T, T2>> expression)
        {
            _builder = builder;
            _expression = expression;
        }

        public void AddError(string message)
        {
            var pos = GetPosition();
            _builder.AddError(pos.row, pos.col, message);
        }

        private (int row, int col) GetPosition()
        {
            throw new Exception();
        }
    }

    public class CollectionHintExecutor<T, T2>
    {
        private HintBuilder<T> _builder;
        private Expression _expression;

        public CollectionHintExecutor(HintBuilder<T> builder, Expression<Func<T, IEnumerable<T2>>> expression)
        {
            _builder = builder;
            _expression = expression;
        }

        public void AddError(string message)
        {
            var pos = GetPosition();
            _builder.AddError(pos.row, pos.col, message);
        }

        private (int row, int col) GetPosition()
        {
            throw new Exception();
        }

        public FieldHintExecutor<T, T2> Item(T2 item)
        {
            //return new FieldHintExecutor<T, T2>(_builder, _expression);
            throw new NotImplementedException();
        }

    }
}
