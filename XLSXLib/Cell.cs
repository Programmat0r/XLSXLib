using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXLib
{
    class Cell
    {
        public String Letter { get; }
        public int Count { get; }

        public Cell(String letter, int count)
        {
            if (letter == null)
                throw new NullReferenceException("There is no EMPTY letter.");

            if (count == 0)
                throw new NullReferenceException("The cell count starts at 1");

            this.Letter = letter;
            this.Count = count;
        }
    }
}
