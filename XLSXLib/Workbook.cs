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

        public String Filename { get; set; }
        public String Name { get; }
        public Worksheet[] Worksheets { get; }

        public Workbook()
        {

        }

        public Workbook(Byte[] file)
        {
            if (File != null)
            {
                this.File = file;
            } else
            {
                throw new FileLoadException("Given file is empty.");
            }
        }

        public Workbook(String filename)
        {
            if (filename != null)
                if (!System.IO.File.Exists(filename))
                    throw new FileNotFoundException("File '" + filename + "' wasn't found.");
           
        }
    }

    


}
