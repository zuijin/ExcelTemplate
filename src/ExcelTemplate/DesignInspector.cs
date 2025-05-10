using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelTemplate.Exceptions;
using ExcelTemplate.Model;
using NPOI.SS.Formula.Functions;

namespace ExcelTemplate
{
    public static class DesignInspector
    {
        /// <summary>
        /// 检验模版设计信息是否符合规范
        /// </summary>
        /// <param name="design"></param>
        /// <exception cref="Exception"></exception>
        public static void Check(TemplateDesign design)
        {
            var fieldMaps = new List<(int row, int col, string fieldPath)>();

            var currentSection = design.BlockSection;
            while (currentSection != null)
            {
                if (IsTableSection(currentSection))
                {
                    var notTable = currentSection.Blocks.Any(a => !(a is TableBlock));
                    if (notTable)
                    {
                        throw new Exception("布局冲突，表单类型和列表类型不能在水平方向上混排");
                    }
                }

                var formFieldMaps = currentSection.Blocks.OfType<ValueBlock>().Select(a => (a.Position.Row, a.Position.Col, a.FieldPath));
                fieldMaps.AddRange(formFieldMaps);

                var tableFieldMaps = currentSection.Blocks.OfType<TableBlock>().SelectMany(a => a.Body.Select(b => (b.Position.Row, b.Position.Col, b.FieldPath)));
                fieldMaps.AddRange(tableFieldMaps);

                currentSection = currentSection.Next;
            }

            fieldMaps = fieldMaps.Distinct().ToList();

            if (design.Usage != TemplateDesignUsage.ImportOnly && fieldMaps.GroupBy(a => (a.row, a.col)).Any(a => a.Count() > 1))
            {
                throw new TemplateDesignException(TemplateDesignExceptionType.PositionConflict,
                    $"发现模版中存在多个字段映射到同一个单元格，这种情况只能使用 {nameof(TemplateDesignUsage.ImportOnly)} 模式，请调整模版设计或者更改 {nameof(design.Usage)} 设置为 {nameof(TemplateDesignUsage.ImportOnly)}");
            }

            if (design.Usage != TemplateDesignUsage.ExportOnly && fieldMaps.GroupBy(a => a.fieldPath).Any(a => a.Count() > 1))
            {
                throw new TemplateDesignException(TemplateDesignExceptionType.FieldConflict,
                    $"发现模版中存在多个单元格映射到同一字段，这种情况只能使用{nameof(TemplateDesignUsage.ExportOnly)}模式，请调整模版设计或者更改 {nameof(design.Usage)} 设置为 {nameof(TemplateDesignUsage.ImportOnly)}");
            }
        }

        /// <summary>
        /// 是否列表区块
        /// </summary>
        /// <param name="page"></param>
        /// <returns></returns>
        public static bool IsTableSection(BlockSection section)
        {
            return section.Blocks.Any(a => a is TableBlock);
        }
    }
}
