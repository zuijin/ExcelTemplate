using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelTemplate.Model
{
    public class BlockSection : ICloneable
    {
        public int BeginRow
        {
            get => Blocks.Min(a => a.Position.Row);
        }

        /// <summary>
        /// 当段落的Block
        /// </summary>
        public List<IBlock> Blocks { get; set; }

        /// <summary>
        /// 下一行
        /// </summary>
        public BlockSection Next { get; set; }

        /// <summary>
        /// 应用位置偏移
        /// </summary>
        /// <param name="rowOffset"></param>
        /// <param name="colOffset"></param>
        public void ApplyOffset(int rowOffset = 0, int colOffset = 0)
        {
            if (rowOffset == 0 && colOffset == 0)
            {
                return;
            }

            foreach (var block in Blocks)
            {
                block.ApplyOffset(rowOffset, colOffset);
            }

            if (this.Next != null)
            {
                this.Next.ApplyOffset(rowOffset, colOffset);
            }
        }

        public object Clone()
        {
            var obj = (BlockSection)this.MemberwiseClone();
            if (this.Blocks != null)
            {
                obj.Blocks = new List<IBlock>();
                foreach (var item in this.Blocks)
                {
                    obj.Blocks.Add((IBlock)item.Clone());
                }
            }

            if (this.Next != null)
            {
                obj.Next = (BlockSection)this.Next.Clone();
            }

            return obj;
        }
    }
}
