using System;

namespace ExcelTemplate.Model
{
    public interface IBlock : ICloneable
    {
        Position Position { get; set; }

        Position MergeTo { get; set; }

        void ApplyOffset(int rowOffset = 0, int colOffset = 0);
    }

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

        public void ApplyOffset(int rowOffset = 0, int colOffset = 0)
        {
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

            return obj;
        }
    }

}
