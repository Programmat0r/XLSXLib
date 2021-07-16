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
        public Cell[] Cells { get; private set; }

        public int AddCell(Cell cell)
        {
            int length = this.Cells.Length + 1;
            Cell[] newCells = new Cell[length];
            this.Cells.CopyTo(newCells, 0);
            this.Cells = newCells;
            this.Cells[length - 1] = cell;

            return length;
        }

        public Row(int rowNumber, Cell[] cells)
        {
            this.RowNumber = rowNumber;
            this.Cells = cells;
        }
    }
}
