using System;
using System.Collections.Generic;
using System.Text;
using NPOI.SS.Formula.Functions;

namespace ExcelTemplate.Model
{
    public class BlockPage
    {
        /// <summary>
        /// 当前行的Block
        /// </summary>
        public List<IBlock> RowBlocks { get; set; }

        /// <summary>
        /// 下一行
        /// </summary>
        public BlockPage Next { get; set; }

        /// <summary>
        /// 应用位置偏移
        /// </summary>
        /// <param name="rowOffset"></param>
        /// <param name="colOffset"></param>
        public void ApplyOffset(int rowOffset = 0, int colOffset = 0)
        {
            foreach (var block in RowBlocks)
            {
                block.Position.ApplyOffset(rowOffset, colOffset);
                if (block.MergeTo != null)
                {
                    block.MergeTo.ApplyOffset(rowOffset, colOffset);
                }
            }

            if (this.Next != null)
            {
                this.Next.ApplyOffset(rowOffset, colOffset);
            }
        }
    }
}
