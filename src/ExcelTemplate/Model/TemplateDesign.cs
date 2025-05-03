using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTemplate.Model
{
    /// <summary>
    /// 模版定义
    /// </summary>
    public class TemplateDesign
    {
        public TemplateDesign() { }
        public TemplateDesign(DesignSourceType sourceType, BlockPage blockPage)
        {
            this.SourceType = sourceType;
            this.BlockPage = blockPage;
        }

        public DesignSourceType SourceType { get; set; }

        public BlockPage BlockPage { get; set; }
    }

    public enum DesignSourceType
    {
        Object = 1,
        File = 2,
    }
}
