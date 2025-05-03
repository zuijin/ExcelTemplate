using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelTemplate.Model;
using NPOI.HSSF.Record.PivotTable;

namespace ExcelTemplate.Helper
{
    public static class MergeHelper
    {

        public static List<TableHeaderBlock> MergeHeader(Position position, List<TypeRawHeader> headerBlocks)
        {
            headerBlocks = headerBlocks.OrderBy(a => a.Block.Position.Col).ToList();
            var maxMergeRows = headerBlocks.Max(a => a.MergeTitles.Length + 1);

            var rootNode = BuildNodeTree(headerBlocks, maxMergeRows);
            HorizontalMerge(rootNode);
            VerticalMerge(rootNode);

            return GetAllBlocks(rootNode, position);
        }

        /// <summary>
        /// 生成树状结构的表头，方便合并
        /// </summary>
        /// <param name="headerBlocks"></param>
        /// <param name="maxHeight"></param>
        /// <returns></returns>
        private static HeaderNode BuildNodeTree(List<TypeRawHeader> headerBlocks, int maxHeight)
        {
            var rootNode = new HeaderNode();
            foreach (var block in headerBlocks)
            {
                var preNode = rootNode;
                var height = 0;
                for (int i = 0; i < block.MergeTitles.Length; i++)
                {
                    var node = new HeaderNode()
                    {
                        Width = 1,
                        Height = 1,
                        Title = block.MergeTitles[i]
                    };

                    preNode.Children.Add(node);
                    preNode = node;
                    height++;
                }

                preNode.Children.Add(new HeaderNode()
                {
                    Width = 1,
                    Height = (maxHeight - height),
                    Title = block.Block.Text?.ToString() ?? "",
                    Block = block.Block,
                });
            }

            return rootNode;
        }

        /// <summary>
        /// 水平合并
        /// </summary>
        /// <param name="node"></param>
        private static void HorizontalMerge(HeaderNode node)
        {
            if (node.Children.Count < 2)
            {
                return;
            }

            HeaderNode preNode = node.Children[0];
            List<HeaderNode> newChildren = new List<HeaderNode>() { preNode };

            for (var i = 1; i < node.Children.Count; i++)
            {
                var currNode = node.Children[i];
                if (currNode.Title == preNode.Title && currNode.Children.Any() && preNode.Children.Any()) //叶子节点不合并
                {
                    preNode.Children.AddRange(currNode.Children);
                    preNode.Width += currNode.Width;
                }
                else
                {
                    newChildren.Add(currNode);
                    preNode = currNode;
                }
            }

            foreach (var child in newChildren)
            {
                HorizontalMerge(child);
            }

            node.Children.Clear();
            node.Children.AddRange(newChildren);
        }

        /// <summary>
        /// 垂直合并
        /// </summary>
        /// <param name="node"></param>
        private static void VerticalMerge(HeaderNode node)
        {
            if (node.Children.Count == 0)
            {
                return;
            }
            else if (node.Children.Count == 1 && node.Title == node.Children[0].Title)
            {
                var child = node.Children[0];
                node.Children.Clear();
                node.Children.AddRange(child.Children);
                node.Height += child.Height;
                node.Block = child.Block;

                VerticalMerge(node);
            }
            else
            {
                foreach (var child in node.Children)
                {
                    VerticalMerge(child);
                }
            }
        }

        /// <summary>
        /// 获取所有 Block
        /// </summary>
        /// <param name="node"></param>
        /// <param name="position"></param>
        /// <param name="heightOffset"></param>
        /// <returns></returns>
        private static List<TableHeaderBlock> GetAllBlocks(HeaderNode node, Position position, int heightOffset = 0, int widthOffset = 0)
        {
            var blocks = new List<TableHeaderBlock>();
            if (node.Width > 0 && node.Height > 0 && node.Children.Count > 0) //非根节点，非叶子节点
            {
                var block = new TableHeaderBlock()
                {
                    Position = position.GetOffset(heightOffset, widthOffset),
                    Text = node.Title,
                };

                if (node.Width > 1 || node.Height > 1)
                {
                    block.MergeTo = block.Position.GetOffset(node.Height - 1, node.Width - 1);
                }

                blocks.Add(block);
            }

            //重设位置信息
            if (node.Block != null) //叶子节点
            {
                var tmp = node.Block;
                var newPosition = position.GetOffset(heightOffset, widthOffset);
                if (tmp.MergeTo != null)
                {
                    var width = tmp.MergeTo.Col - tmp.Position.Col;
                    var height = tmp.MergeTo.Row - tmp.Position.Row;
                    tmp.MergeTo = newPosition.GetOffset(height, width);
                }

                tmp.Position = newPosition;
                blocks.Add(tmp);
            }

            int totalWidth = 0;
            foreach (var child in node.Children)
            {
                var childBlocks = GetAllBlocks(child, position, heightOffset + node.Height, totalWidth + widthOffset);
                blocks.AddRange(childBlocks);
                totalWidth += child.Width;
            }

            return blocks;
        }


        private class HeaderNode
        {
            public int Width { get; set; }

            public int Height { get; set; }

            public string Title { get; set; }

            /// <summary>
            /// 叶子节点才会有值，挂载原始的区块定义
            /// </summary>
            public TableHeaderBlock Block { get; set; }

            public List<HeaderNode> Children { get; set; } = new List<HeaderNode>();
        }
    }


}
