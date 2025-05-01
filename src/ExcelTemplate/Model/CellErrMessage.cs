using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTemplate.Model
{
    public class CellErrMessage
    {
        public int RowIndex { get; set; }

        public int ColIndex { get; set; }

        public string Message { get; set; }
    }
}
