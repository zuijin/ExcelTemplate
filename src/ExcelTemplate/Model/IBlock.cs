using System;

namespace ExcelTemplate.Model
{
    public interface IBlock : ICloneable
    {
        Position Position { get; set; }

        Position MergeTo { get; set; }

        void ApplyOffset(int rowOffset = 0, int colOffset = 0);
    }
}
