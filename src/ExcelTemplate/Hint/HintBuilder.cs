using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using ExcelTemplate.Helper;
using ExcelTemplate.Model;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XWPF.UserModel;

namespace ExcelTemplate.Hint
{
    public class HintBuilder<T>
    {
        private T _data;
        private List<CellException> _exceptions;
        private IWorkbook _workbook;
        private TemplateDesign _design;

        public T Data => _data;
        internal Dictionary<string, Position> FieldDic = new Dictionary<string, Position>();
        internal Dictionary<object, int> ElemetDic = new Dictionary<object, int>();

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
            var current = _design.BlockSection;
            while (current != null)
            {
                foreach (var block in current.Blocks.OfType<ValueBlock>())
                {
                    FieldDic.Add(block.FieldPath, block.Position);
                }

                foreach (var table in current.Blocks.OfType<TableBlock>())
                {
                    var list = (IEnumerable)ObjectHelper.GetObjectValue(_data, table.TableName);
                    var index = 0;
                    foreach (var element in list)
                    {
                        var firstRow = table.Body.First().Position.Row;
                        foreach (var body in table.Body)
                        {
                            var fieldPath = body.FieldPath.Substring(body.FieldPath.IndexOf('.') + 1);
                            var dataPath = $"{table.TableName}.{index}.{fieldPath}";
                            FieldDic.Add(dataPath, (firstRow + index, body.Position.Col));
                        }

                        ElemetDic.Add(element, index);
                        index++;
                    }
                }

                current = current.Next;
            }
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
