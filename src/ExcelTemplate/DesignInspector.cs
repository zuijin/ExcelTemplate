using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelTemplate.Model;

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

                currentSection = currentSection.Next;
            }

            //TODO: 检查是否有多个字段映射到同一个单元格的情况
            //TODO: 检查是否有多个单元格映射到同一个字段的情况
            //TODO: 检查是否有某个字段是只读状态，包括只有get访问器的情况 
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
