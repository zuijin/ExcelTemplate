using ExcelTemplate.Style;

namespace ExcelTemplate.Model
{
    public class BlockBase : IBlock
    {
        /// <summary>
        /// 位置
        /// </summary>
        public Position Position { get; set; }
        /// <summary>
        /// 合并到哪个位置
        /// </summary>
        public Position MergeTo { get; set; }
        /// <summary>
        /// 样式
        /// </summary>
        public IStyle Style { get; set; }

        public virtual void ApplyOffset(int rowOffset = 0, int colOffset = 0)
        {
            if (rowOffset == 0 && colOffset == 0)
            {
                return;
            }

            this.Position.ApplyOffset(rowOffset, colOffset);
            if (this.MergeTo != null)
            {
                this.MergeTo.ApplyOffset(rowOffset, colOffset);
            }
        }

        public virtual object Clone()
        {
            var obj = (BlockBase)this.MemberwiseClone();
            obj.Position = (Position)this.Position.Clone();

            if (this.MergeTo != null)
            {
                obj.MergeTo = (Position)this.MergeTo.Clone();
            }

            if (this.Style != null)
            {
                obj.Style = (IStyle)this.Style.Clone();
            }

            return obj;
        }
    }

}
