using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace XLSXLib
{
    public class Worksheet
    {

        public String Title { get; }

        public String SystemName { get; private set; }
        public Row[] Rows { get; }

        public Worksheet(String title, Row[] rows)
        {
            this.Title = title;
            this.Rows = rows;
        }

        public Worksheet(String systemName, String title, Row[] rows)
        {
            this.SystemName = systemName;
            this.Title = title;
            this.Rows = rows;
        }
    }

}
