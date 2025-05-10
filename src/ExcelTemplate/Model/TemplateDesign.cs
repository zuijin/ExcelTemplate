using System;
using System.Collections.Generic;
using ExcelTemplate.Style;

namespace ExcelTemplate.Model
{
    /// <summary>
    /// 模版定义
    /// </summary>
    public class TemplateDesign : ICloneable
    {
        public TemplateDesign() { }
        public TemplateDesign(TemplateDesignSourceType sourceType, BlockSection BlockSection)
        {
            this.SourceType = sourceType;
            this.BlockSection = BlockSection;
        }

        /// <summary>
        /// 模版设计来源
        /// </summary>
        public TemplateDesignSourceType SourceType { get; private set; }

        /// <summary>
        /// 模版使用方式
        /// </summary>
        public TemplateDesignUsage Usage { get; set; } = TemplateDesignUsage.TwoWay;

        /// <summary>
        /// 区块定义
        /// </summary>
        public BlockSection BlockSection { get; set; }

        public object Clone()
        {
            var obj = (TemplateDesign)this.MemberwiseClone();
            if (this.BlockSection != null)
            {
                obj.BlockSection = (BlockSection)this.BlockSection.Clone();
            }

            return obj;
        }
    }

    public enum TemplateDesignSourceType
    {
        /// <summary>
        /// 类型
        /// </summary>
        Type = 1,
        /// <summary>
        /// Excel文件
        /// </summary>
        File = 2,
    }

    public enum TemplateDesignUsage
    {
        /// <summary>
        /// 双向
        /// </summary>
        TwoWay = 0,
        /// <summary>
        /// 只导入
        /// </summary>
        ImportOnly = 1,
        /// <summary>
        /// 只导出
        /// </summary>
        ExportOnly = 2,
    }
}
