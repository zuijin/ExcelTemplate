﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using ExcelTemplate.Exceptions;
using ExcelTemplate.Extensions;
using ExcelTemplate.Helper;
using ExcelTemplate.Model;
using ExcelTemplate.Style;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;

namespace ExcelTemplate.Hint
{
    public class HintBuilder<T>
    {
        private T _data;
        private List<CellHintMessage> _messages;
        private IWorkbook _workbook;
        private TemplateDesign _design;
        private Dictionary<string, Position> _fieldPositionDic;
        private Dictionary<object, int> _elemetIndexDic;
        private string _messageBgColor;

        public T Data => _data;
        public IWorkbook Workbook => _workbook;
        internal Dictionary<string, Position> FieldPositionDic => _fieldPositionDic;
        internal Dictionary<object, int> ElemetIndexDic => _elemetIndexDic;

        public HintBuilder(TemplateDesign design, IWorkbook workbook, T data, List<CellHintMessage> messages)
        {
            _messages = new List<CellHintMessage>();
            _messages.AddRange(messages);
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

        /// <summary>
        /// 设置提示单元格的背景颜色，十六进制RGB格式 FFF 或 FFFFFF，如果传 null 则不更改背景颜色
        /// </summary>
        /// <param name="bgColor"></param>
        /// <exception cref="Exception"></exception>
        public void SetMessageBgColor(string bgColor)
        {
            if (string.IsNullOrWhiteSpace(bgColor))
            {
                _messageBgColor = null;
                return;
            }

            if (!ETStyleUtil.IsRgbHex(bgColor))
            {
                throw new Exception("颜色格式错误，不符合十六进制 FFF 或 FFFFFF 的格式");
            }

            _messageBgColor = bgColor;
        }

        /// <summary>
        /// 添加提示信息
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="message"></param>
        public void AddMessage(int row, int col, string message)
        {
            _messages.Add(new CellHintMessage(row, col, message));
        }

        public FieldHintExp<T, TField> For<TField>(Expression<Func<T, TField>> expression)
        {
            return new FieldHintExp<T, TField>(this, expression);
        }

        /// <summary>
        /// 生成提示Excel
        /// </summary>
        /// <returns></returns>
        public IWorkbook BuildExcel()
        {
            var newWorkbook = _workbook.Copy();
            var sheet = newWorkbook.GetSheetAt(0);
            var drawing = sheet.CreateDrawingPatriarch();
            var helper = newWorkbook.GetCreationHelper();

            foreach (var group in _messages.GroupBy(a => (a.Position.Row, a.Position.Col)))
            {
                var position = group.Key;
                var message = string.Join(Environment.NewLine, group.Select(a => a.Message));
                var cell = sheet.GetCell(position.Row, position.Col);

                //设置批注
                var anchor = helper.CreateClientAnchor();
                var comment = drawing.CreateCellComment(anchor);
                comment.String = helper.CreateRichTextString(message);
                cell.CellComment = comment;

                //设置背景颜色
                if (!string.IsNullOrWhiteSpace(_messageBgColor))
                {
                    var newStyle = newWorkbook.CreateCellStyle();
                    newStyle.CloneStyleFrom(cell.CellStyle);
                    newStyle.FillPattern = FillPattern.SolidForeground;

                    if (newStyle is XSSFCellStyle)
                    {
                        ((XSSFCellStyle)newStyle).SetFillForegroundColor(ETStyleUtil.GetXSSFColor(_messageBgColor));
                    }
                    else
                    {
                        newStyle.FillForegroundColor = ETStyleUtil.GetHSSFColor(newWorkbook, _messageBgColor);
                    }

                    cell.CellStyle = newStyle;
                }
            }

            return newWorkbook;
        }

        /// <summary>
        /// 生成提示
        /// </summary>
        /// <returns></returns>
        public string BuildMessage()
        {
            if (!_messages.Any())
            {
                return string.Empty;
            }

            StringBuilder sb = new StringBuilder(500);
            foreach (var ex in _messages.OrderBy(a => a.Position.Row))
            {
                sb.AppendLine(ex.Message);
            }

            return sb.ToString();
        }

        /// <summary>
        /// 生成提示文件
        /// </summary>
        /// <returns></returns>
        public byte[] BuildFile()
        {
            using (var ms = new MemoryStream())
            {
                var book = BuildExcel();
                book.Write(ms);

                return ms.ToArray();
            }
        }
    }
}
