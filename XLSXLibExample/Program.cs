using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
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

            workbook.Worksheets[0].Rows[0].Cells[0].Value = "2";
            workbook.Worksheets[0].Rows[0].AddCell(new Cell("B1", "1"));

            workbook.Save();

            Console.WriteLine(workbook.Worksheets[0].Title);
            foreach (var row in workbook.Worksheets[0].Rows)
            {
                foreach (var cell in row.Cells)
                {
                    Console.Write(cell.Location);
                   
                }
                Console.WriteLine("");
            }
            Console.ReadKey();
        }
    }
}
