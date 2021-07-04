using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXLib
{
   public class Row
    {
        public int RowNumber { get; }
        public Cell[] Cells { get; }

        public Row(int rowNumber, Cell[] cells)
        {
            this.RowNumber = rowNumber;
            this.Cells = cells;
        }
    }
}
