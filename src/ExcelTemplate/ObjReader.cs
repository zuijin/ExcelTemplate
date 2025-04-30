using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using ExcelTemplate.Attributes;
using ExcelTemplate.Helper;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;

namespace ExcelTemplate
{
    public class ObjReader<T> where T : new()
    {
        IWorkbook _workbook;
        List<FieldInfo> _fieldInfos = new List<FieldInfo>();

        T _data = new T();
        List<CellException> _exceptionList = new List<CellException>();

        /// <summary>
        /// 
        /// </summary>
        public ObjReader(IWorkbook workbook)
        {
            _workbook = workbook;
            ReadFieldInfos();
            ReadData();
        }

        /// <summary>
        /// 从文件中生成 ExcelWrap
        /// </summary>
        public static ObjReader<T> FromFile(Stream file, string fileName)
        {
            MemoryStream ms = new MemoryStream();
            file.CopyTo(ms);
            ms.Seek(0, SeekOrigin.Begin);

            IWorkbook book;

            string fileExt = Path.GetExtension(fileName).ToLower();
            if (fileExt == ".xlsx")
            {
                book = new XSSFWorkbook(ms);
            }
            else if (fileExt == ".xls")
            {
                book = new HSSFWorkbook(ms);
            }
            else
            {
                throw new Exception("不支持的文件类型");
            }

            return new ObjReader<T>(book);
        }

        /// <summary>
        /// 是否存在转换错误
        /// </summary>
        /// <returns></returns>
        public bool IsError => _exceptionList.Any();

        /// <summary>
        /// 获取数据集合
        /// </summary>
        /// <returns></returns>
        public T GetData()
        {
            if (IsError)
            {
                throw _exceptionList[0];
            }

            return _data;
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

            foreach (var ex in _exceptionList)
            {
                var cell = sheet.GetRow(ex.Row).GetCell(ex.Column);
                var anchor = helper.CreateClientAnchor();
                var comment = drawing.CreateCellComment(anchor);
                comment.String = helper.CreateRichTextString(ex.Message);
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
            if (!IsError)
            {
                return string.Empty;
            }

            StringBuilder sb = new StringBuilder(500);
            foreach (var ex in _exceptionList.OrderBy(a => a.Row))
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

        /// <summary>
        /// 转换为数据列表
        /// </summary>
        private void ReadData()
        {
            _exceptionList.Clear();
            var sheet = _workbook.GetSheetAt(0);

            foreach (var field in _fieldInfos.Where(a => a.ValuePosition != null))
            {
                var pos = field.ValuePosition;
                var row = sheet.GetRow(pos.Row);
                if (row == null) continue;

                var cell = row.GetCell(pos.Col, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                if (cell == null) continue;

                var type = field.Property.PropertyType;
                try
                {
                    var val = GetCellValue(cell, field.Property.PropertyType);
                    field.Property.SetValue(_data, val);
                }
                catch (Exception ex)
                {
                    throw new CellException(field.RowIndex, field.ColIndex, $"行{field.RowIndex + 1}列{field.ColIndex + 1}，数据输入错误", ex);
                }
            }
        }

        object? GetCellValue(ICell cell, Type type)
        {
            object? val = type.IsValueType ? Activator.CreateInstance(type) : null;
            if (cell == null)
            {
                return val;
            }

            switch (cell.CellType)
            {
                case CellType.String:
                    val = cell.StringCellValue;
                    break;
                case CellType.Numeric:
                    val = DateUtil.IsCellDateFormatted(cell) ? (object?)cell.DateCellValue : cell.NumericCellValue;
                    break;
                case CellType.Boolean:
                    val = cell.BooleanCellValue;
                    break;
                case CellType.Blank:
                    break;
                default:
                    val = cell.ToString();
                    break;
            }

            return System.Convert.ChangeType(val, type);
        }


        /// <summary>
        /// 判断是否可 null 类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        private bool IsNullableType(Type type)
        {
            return !type.IsValueType
                || type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>);
        }

        /// <summary>
        /// 生成列信息
        /// </summary>
        /// <returns></returns>
        private void ReadFieldInfos()
        {
            var props = typeof(T).GetProperties();
            foreach (PropertyInfo prop in props)
            {
                var info = new FieldInfo()
                {
                    Property = prop,
                };

                var titleAttr = prop.GetCustomAttribute<TitleAttribute>();
                if (titleAttr != null)
                {
                    info.Title = titleAttr.Title;
                    info.TitlePosition = new Position(titleAttr.Position);

                    if (!string.IsNullOrEmpty(titleAttr.MergeTo))
                    {
                        info.TitleMergeTo = new Position(titleAttr.MergeTo);
                    }
                }

                var valueAttr = prop.GetCustomAttribute<ValueAttribute>();
                if (valueAttr != null)
                {
                    info.ValuePosition = new Position(valueAttr.Position);
                    if (!string.IsNullOrWhiteSpace(valueAttr.MergeTo))
                    {
                        info.ValueMergeTo = new Position(valueAttr.MergeTo);
                    }
                }

                var mergeAttr = prop.GetCustomAttribute<MergeAttribute>();
                if (mergeAttr != null)
                {
                    info.HeaderMerge = mergeAttr.Titles;
                }

                _fieldInfos.Add(info);
            }
        }

        public List<Block> DefinitionAnalysis(Type type, List<PropertyInfo>? parents = null)
        {
            var blocks = new List<Block>();
            var props = type.GetProperties();
            foreach (PropertyInfo prop in props)
            {
                var propType = prop.PropertyType;
                if (TypeHelper.IsSimpleType(propType))
                {
                    blocks.AddRange(GetSimpleTypeBlocks(prop, parents));
                }
                else if (IsSpecialCollectionType(propType))
                {
                    blocks.AddRange(GetCollectionTypeBlocks(prop, parents));
                }
                else
                {
                    throw new Exception($"暂不支持复杂类型：{prop.Name}");
                }
            }

            return blocks;
        }


        /// <summary>
        /// 判断是否支持的集合类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static bool IsSpecialCollectionType(Type type)
        {
            if (!TypeHelper.IsCollectionType(type))
                return false;

            // 获取集合元素的类型
            Type elementType = TypeHelper.GetCollectionElementType(type);

            // 判断元素类型是否为简单类型
            return !TypeHelper.IsSimpleType(elementType) && !TypeHelper.IsCollectionType(elementType);
        }

        /// <summary>
        /// 获取简单类型Block
        /// </summary>
        /// <param name="prop"></param>
        /// <param name="parent"></param>
        /// <returns></returns>
        public List<Block> GetSimpleTypeBlocks(PropertyInfo prop, List<PropertyInfo>? parents = null)
        {
            var blocks = new List<Block>();
            var titleAttr = prop.GetCustomAttribute<TitleAttribute>();
            if (titleAttr != null)
            {
                blocks.Add(new Block()
                {
                    BlockType = BlockType.Constant,
                    Position = titleAttr.Position,
                    Value = titleAttr.Title,
                    MergeTo = titleAttr.MergeTo,
                });
            }

            var valueAttr = prop.GetCustomAttribute<ValueAttribute>();
            if (valueAttr != null)
            {
                blocks.Add(new Block()
                {
                    BlockType = BlockType.Variable,
                    Position = valueAttr.Position,
                    MergeTo = valueAttr.MergeTo,
                    ValuePath = PathCombine(parents?.Select(a => a.Name), prop.Name)
                });
            }

            return blocks;
        }

        /// <summary>
        /// 获取集合类型Block
        /// </summary>
        /// <param name="prop"></param>
        /// <param name="parent"></param>
        /// <returns></returns>
        public List<Block> GetCollectionTypeBlocks(PropertyInfo prop, List<PropertyInfo>? parents = null)
        {
            var blocks = new List<Block>();
            var valueAttr = prop.GetCustomAttribute<ValueAttribute>();
            if (valueAttr == null)
            {
                return blocks;
            }

            var elementType = TypeHelper.GetCollectionElementType(prop.PropertyType);
            var subProps = elementType.GetProperties();
            var tmpList = new List<(Block block, string[] mergeTitles)>();

            foreach (var subProp in subProps)
            {
                var path = PathCombine(parents?.Select(a => a.Name), prop.Name);
                var propType = subProp.PropertyType;
                if (!TypeHelper.IsSimpleType(propType))
                {
                    throw new Exception($"集合内不支持任何复杂类型：{path}");
                }

                var titleAttr = prop.GetCustomAttribute<TitleAttribute>();
                if (titleAttr != null)
                {
                    var block = new Block()
                    {
                        BlockType = BlockType.Array,
                        Position = titleAttr.Position,
                        ValuePath = path,
                    };

                    string[] mergeTitles = [];
                    var mergeAttr = prop.GetCustomAttribute<MergeAttribute>();
                    if (mergeAttr != null)
                    {
                        mergeTitles = mergeAttr.Titles;
                    }

                    tmpList.Add((block, mergeTitles));
                }
            }

            var maxMergeRows = tmpList.Max(a => a.mergeTitles.Length);
            var firstCol = new Position(valueAttr.Position).Col;
            foreach (var item in tmpList)
            {

            }

            return blocks;
        }

        private string PathCombine(params string?[] paths)
        {
            return string.Join(".", paths.Where(a => !string.IsNullOrWhiteSpace(a)));
        }

        private string PathCombine(IEnumerable<string?>? paths1, params string?[] paths2)
        {
            List<string?> paths = new List<string?>();
            if (paths1 != null)
            {
                paths.AddRange(paths1);
            }

            paths.AddRange(paths2);

            return string.Join(".", paths.Where(a => !string.IsNullOrWhiteSpace(a)));
        }

        /// <summary>
        /// 添加错误提示
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="colName"></param>
        /// <param name="message"></param>
        /// <exception cref="NotImplementedException"></exception>
        public void AddError(string colName, string message)
        {
            var col = _fieldInfos.Find(a => a.FieldName == colName || a.Title == colName);
            int colIndex = col?.ColIndex ?? -1;
            int rowIndexd = col?.RowIndex ?? -1;

            _exceptionList.Add(new CellException(rowIndexd, colIndex, message));
        }
    }
}
