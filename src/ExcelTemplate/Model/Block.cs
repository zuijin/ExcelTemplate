namespace ExcelTemplate.Model
{
    public class Block
    {
        /// <summary>
        /// 类型
        /// </summary>
        public BlockType BlockType { get; set; }

        public object Value { get; set; }

        public string ValuePath { get; set; }

        public Position Position { get; set; }

        public Position MergeTo { get; set; }
    }
}
