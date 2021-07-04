using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXLib
{
    public class Workbook
    {
        public Byte[] File { get; set; }
        public String Name { get; }
        public Worksheet[] Worksheets { get; }

        public Workbook()
        {

        }

        public Workbook(Byte[] File)
        {
            if (File != null)
            {
                this.File = File;
            } else
            {
                throw new FileLoadException("Given file is empty.");
            }
        }
    }

    


}
