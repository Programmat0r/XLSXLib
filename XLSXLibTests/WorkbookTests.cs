using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLSXLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace XLSXLib.Tests
{
    [TestClass()]
    public class WorkbookTests
    {
        [TestMethod()]
        public void Workbook_Wrong_File()
        {
            Assert.ThrowsException<FileNotFoundException>(() => new Workbook("Example/Testfile.xls"));
        }

        [TestMethod()]
        public void Workbook_Corrupted_File()
        {
            var workbook = new Workbook("Example/Corrupt.xlsx");

            Assert.ThrowsException<InvalidDataException>(() => workbook.Load());
        }
    }
}