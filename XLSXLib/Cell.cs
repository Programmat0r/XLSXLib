using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXLib
{
    public class Cell
    {
        public String Location { get; }
        public String Value { get; set; }

        public Cell(String location, String value)
        {
            if (location == null)
                throw new NullReferenceException("The location can't be empty");

            this.Location = location;
            this.Value = value;
        }
    }
}
