using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using ExcelTemplate.Extensions;
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
        private Dictionary<string, Position> _fieldPositionDic;
        private Dictionary<object, int> _elemetIndexDic;

        public T Data => _data;
        internal Dictionary<string, Position> FieldPositionDic => _fieldPositionDic;
        internal Dictionary<object, int> ElemetIndexDic => _elemetIndexDic;

        public HintBuilder(TemplateDesign design, IWorkbook workbook, T data, List<CellException> exceptions)
        {
            _exceptions = new List<CellException>();
            _exceptions.AddRange(exceptions);
            _design = design;
            _workbook = workbook;
            _data = data;
            _fieldPositionDic = new Dictionary<string, Position>();
            _elemetIndexDic = new Dictionary<object, int>();

            InitDic();
        }

        private void InitDic()
        {
            var current = _design.BlockSection;
            while (current != null)
            {
                foreach (var block in current.Blocks.OfType<ValueBlock>())
                {
                    FieldPositionDic.Add(block.FieldPath, block.Position);
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
                            FieldPositionDic.Add(dataPath, (firstRow + index, body.Position.Col));
                        }

                        ElemetIndexDic.Add(element, index);
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

        /// <summary>
        /// 是否存在异常
        /// </summary>
        /// <returns></returns>
        public bool HasError()
        {
            return _exceptions.Any();
        }

        /// <summary>
        /// 生成错误提示Excel
        /// </summary>
        /// <returns></returns>
        public IWorkbook BuildErrorExcel()
        {
            var newWorkbook = _workbook.Copy();
            var sheet = newWorkbook.GetSheetAt(0);
            var drawing = sheet.CreateDrawingPatriarch();
            var helper = _workbook.GetCreationHelper();

            foreach (var group in _exceptions.GroupBy(a => (a.Position.Row, a.Position.Col)))
            {
                var position = group.Key;
                var message = string.Join(Environment.NewLine, group.Select(a => a.Message));

                var cell = sheet.GetCell(position.Row, position.Col);
                var anchor = helper.CreateClientAnchor();
                var comment = drawing.CreateCellComment(anchor);
                comment.String = helper.CreateRichTextString(message);
                cell.CellComment = comment;
            }

            return newWorkbook;
        }

        /// <summary>
        /// 生成错误提示
        /// </summary>
        /// <returns></returns>
        public string BuildErrorMessage()
        {
            if (!_exceptions.Any())
            {
                return string.Empty;
            }

            StringBuilder sb = new StringBuilder(500);
            foreach (var ex in _exceptions.OrderBy(a => a.Position.Row))
            {
                sb.AppendLine(ex.Message);
            }

            return sb.ToString();
        }

        /// <summary>
        /// 生成错误提示文件
        /// </summary>
        /// <returns></returns>
        public byte[] BuildErrorFile()
        {
            using (var ms = new MemoryStream())
            {
                var book = BuildErrorExcel();
                book.Write(ms);

                return ms.ToArray();
            }
        }
    }
}
