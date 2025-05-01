
//using System;
//using System.Collections.Generic;
//using System.ComponentModel;
//using System.IO;
//using System.Linq;
//using System.Reflection;
//using ExcelTemplate.Model;
//using NPOI.HSSF.UserModel;
//using NPOI.SS.Formula.Functions;
//using NPOI.SS.UserModel;
//using NPOI.XSSF.UserModel;

//namespace ExcelTemplate
//{
//    /// <summary>
//    /// 
//    /// </summary>
//    /// <typeparam name="T"></typeparam>
//    public class ListReader
//    {
//        ISheet _sheet;
//        int _titleIndex;
//        Type _type;

//        List<FieldInfo> _columnInfos = new List<FieldInfo>();

//        List<object> _dataList = new List<object>();
//        List<CellException> _exceptionList = new List<CellException>();

//        /// <summary>
//        /// 获取数据集合
//        /// </summary>
//        public ListReader(ISheet sheet, int titleIndex, Type type)
//        {
//            _sheet = sheet;
//            _titleIndex = titleIndex;
//            _type = type;

//            GetColumnInfos();
//            ConvertToList();
//        }

//        /// <summary>
//        /// 从文件中生成 ExcelWrap
//        /// </summary>
//        public static ExcelListWrap FromFile(Stream file, string fileName, int titleIndex, Type type)
//        {
//            MemoryStream ms = new MemoryStream();
//            file.CopyTo(ms);
//            ms.Seek(0, SeekOrigin.Begin);

//            IWorkbook book;

//            string fileExt = Path.GetExtension(fileName).ToLower();
//            if (fileExt == ".xlsx")
//            {
//                book = new XSSFWorkbook(ms);
//            }
//            else if (fileExt == ".xls")
//            {
//                book = new HSSFWorkbook(ms);
//            }
//            else
//            {
//                throw new Exception("不支持的文件类型");
//            }

//            return new ExcelListWrap(book, titleIndex, type);
//        }

//        /// <summary>
//        /// 是否存在转换错误
//        /// </summary>
//        /// <returns></returns>
//        public bool IsError => _exceptionList.Any();

//        /// <summary>
//        /// 获取数据集合
//        /// </summary>
//        /// <returns></returns>
//        public List<object> GetData()
//        {
//            if (IsError)
//            {
//                throw _exceptionList[0];
//            }

//            return _dataList;
//        }

//        /// <summary>
//        /// 获取数据集合
//        /// </summary>
//        /// <typeparam name="T"></typeparam>
//        /// <returns></returns>
//        public List<T> GetData<T>()
//        {
//            var list = GetData();
//            return list.Select(a => (T)a).ToList();
//        }

//        /// <summary>
//        /// 生成错误提示Excel
//        /// </summary>
//        /// <returns></returns>
//        public IWorkbook BuildErrorExcel()
//        {
//            var newWorkbook = Copy(_workbook);
//            var sheet = newWorkbook.GetSheetAt(0);

//            var titleRow = sheet.GetRow(_titleIndex);
//            var errColIndex = _columnInfos.Max(a => a.ColIndex) + 1;

//            titleRow.CreateCell(errColIndex).SetCellValue("错误原因");

//            foreach (var exGroup in _exceptionList.GroupBy(a => a.Row))
//            {
//                var row = sheet.GetRow(exGroup.Key);
//                var message = string.Join("；", exGroup.Select(a => a.Message));
//                row.CreateCell(errColIndex).SetCellValue(message);
//            }

//            return newWorkbook;
//        }

//        /// <summary>
//        /// 生成错误提示
//        /// </summary>
//        /// <returns></returns>
//        public string BuildErrorMessage()
//        {
//            if (!IsError)
//            {
//                return string.Empty;
//            }

//            return _exceptionList
//                 .OrderBy(a => a.Row)
//                 .Select(a => a.Message)
//                 .Aggregate((a, b) => $"{a};\n{b}");
//        }

//        /// <summary>
//        /// 生成错误提示文件
//        /// </summary>
//        /// <returns></returns>
//        public byte[] BuildErrorFile()
//        {
//            using (var ms = new MemoryStream())
//            {
//                var book = BuildErrorExcel();
//                book.Write(ms);

//                return ms.ToArray();
//            }
//        }

//        /// <summary>
//        /// 复制Excel
//        /// </summary>
//        /// <param name="workbook"></param>
//        /// <returns></returns>
//        private IWorkbook Copy(IWorkbook workbook)
//        {
//            MemoryStream ms = new MemoryStream();
//            workbook.Write(ms);

//            ms.Seek(0, SeekOrigin.Begin);

//            if (workbook is XSSFWorkbook)
//            {
//                return new XSSFWorkbook(ms);
//            }
//            else
//            {
//                return new HSSFWorkbook(ms);
//            }
//        }

//        /// <summary>
//        /// 转换为数据列表
//        /// </summary>
//        private void ConvertToList()
//        {
//            _dataList.Clear();
//            _exceptionList.Clear();

//            var sheet = _workbook.GetSheetAt(0);

//            for (int index = 0, count = sheet.LastRowNum; index <= count; index++)
//            {
//                if (index <= _titleIndex)
//                {
//                    continue;
//                }

//                try
//                {
//                    IRow row = sheet.GetRow(index);
//                    if (!IsEmptyRow(row))
//                    {
//                        var tmp = ConvertRow(row, index);
//                        _dataList.Add(tmp);
//                    }
//                }
//                catch (CellException ex)
//                {
//                    _exceptionList.Add(ex);
//                }
//            }
//        }

//        /// <summary>
//        /// 根据反射 解析Excel行信息
//        /// </summary>
//        /// <returns></returns>
//        private object? ConvertRow(IRow row, int rowIndex)
//        {
//            object? model = Activator.CreateInstance(_type);

//            foreach (var col in _columnInfos.Where(a => a.Property != null))
//            {
//                object? val = null;
//                try
//                {
//                    var cell = row.GetCell(col.ColIndex, MissingCellPolicy.CREATE_NULL_AS_BLANK);
//                    val = GetCellValue(cell, col.Property.PropertyType);
//                }
//                catch (Exception ex)
//                {
//                    throw new CellException(rowIndex, col.ColIndex, $"行{rowIndex + 1}列{col.ColIndex + 1}，数据输入错误", ex);
//                }

//                col.Property.SetValue(model, val);
//            }

//            return model;
//        }

//        object? GetCellValue(ICell cell, Type type)
//        {
//            object? val = type.IsValueType ? Activator.CreateInstance(type) : null;
//            if (cell == null)
//            {
//                return val;
//            }

//            switch (cell.CellType)
//            {
//                case CellType.String:
//                    val = cell.StringCellValue;
//                    break;
//                case CellType.Numeric:
//                    val = DateUtil.IsCellDateFormatted(cell) ? cell.DateCellValue : cell.NumericCellValue;
//                    break;
//                case CellType.Boolean:
//                    val = cell.BooleanCellValue;
//                    break;
//                case CellType.Blank:
//                    break;
//                default:
//                    val = cell.ToString();
//                    break;
//            }

//            return Convert.ChangeType(val, type);
//        }

//        /// <summary>
//        /// 判断是否可 null 类型
//        /// </summary>
//        /// <param name="type"></param>
//        /// <returns></returns>
//        private bool IsNullableType(Type type)
//        {
//            return !type.IsValueType
//                || type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>);
//        }

//        /// <summary>
//        /// 生成列信息
//        /// </summary>
//        /// <returns></returns>
//        private void GetColumnInfos()
//        {
//            var sheet = _workbook.GetSheetAt(0);
//            var props = _type.GetProperties();
//            foreach (var prop in props)
//            {
//                var attr = prop.GetCustomAttribute<FieldAttribute>();
//                if (attr != null)
//                {
//                    _columnInfos.Add(new FieldInfo()
//                    {
//                        Title = attr.Title,
                      
//                        FieldName = prop?.Name,
//                        Property = prop,
//                    });
//                }
//            }
//        }

//        /// <summary>
//        /// 判断是否空行
//        /// </summary>
//        /// <param name="row"></param>
//        /// <returns></returns>
//        private static bool IsEmptyRow(IRow row)
//        {
//            if (row == null)
//            {
//                return true;
//            }

//            var enumerator = row.GetEnumerator();
//            while (enumerator.MoveNext())
//            {
//                if (enumerator.Current != null && !string.IsNullOrWhiteSpace(enumerator.Current.ToString()))
//                {
//                    return false;
//                }
//            }

//            return true;
//        }

//        /// <summary>
//        /// 添加错误提示
//        /// </summary>
//        /// <param name="rowIndex"></param>
//        /// <param name="colName"></param>
//        /// <param name="message"></param>
//        /// <exception cref="NotImplementedException"></exception>
//        public void AddError(int rowIndex, string colName, string message)
//        {
//            var col = _columnInfos.Find(a => a.FieldName == colName || a.Title == colName);
//            int colIndex = col?.ColIndex ?? -1;

//            _exceptionList.Add(new CellException(rowIndex, colIndex, message));
//        }
//    }
//}
