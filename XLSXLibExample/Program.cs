using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXLib;

namespace XLSXLibExample
{
    class Program
    {
        static void Main(string[] args)
        {
            var workbook = new Workbook("Example/Testfile.xlsx");
            workbook.Load();

        }
    }
}
