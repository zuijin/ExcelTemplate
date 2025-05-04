namespace ExcelTemplate.Model
{
    /// <summary>
    /// 模版定义
    /// </summary>
    public class TemplateDesign
    {
        public TemplateDesign() { }
        public TemplateDesign(DesignSourceType sourceType, BlockSection BlockSection)
        {
            this.SourceType = sourceType;
            this.BlockSection = BlockSection;
        }

        public DesignSourceType SourceType { get; set; }

        public BlockSection BlockSection { get; set; }
    }

    public enum DesignSourceType
    {
        Object = 1,
        File = 2,
    }
}
