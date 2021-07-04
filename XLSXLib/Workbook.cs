using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace XLSXLib
{
    public class Workbook
    {
        public Byte[] File { get; set; }

        public String Filename { get; set; }
        public String Name { get; }
        public Worksheet[] Worksheets { get; private set; }


        public Workbook()
        {

        }

        public void Load()
        {
            var doc = new XmlDocument();
       
            if (this.Filename != null)
            {


                using (var data = System.IO.File.OpenRead(this.Filename))
                    
                using (var zip = new ZipArchive(data, ZipArchiveMode.Read))
                 {
                    var workbook = zip.GetEntry("xl/workbook.xml");

                    var docWorksheets = new List<String>();

                    using (var stream = workbook.Open())
                    {
                        var sr = new StreamReader(stream);
                        doc.LoadXml(sr.ReadToEnd());

                        var sheets = doc.GetElementsByTagName("sheet");



                        foreach (XmlNode sheet in sheets)
                        {
                            docWorksheets.Add(sheet.Attributes.GetNamedItem("name").Value);
                        }
                    }

                    var worksheetList = new List<Worksheet>();
                    int worksheetCounter = 0;

                        foreach (var entry in zip.Entries)
                        {
                        if(entry.FullName.Contains("xl/worksheets/"))
                        {

                            using (var stream = entry.Open())
                            {
                                var sr = new StreamReader(stream);
                                doc.LoadXml(sr.ReadToEnd());

                                var docRows = doc.GetElementsByTagName("row");

                                var rows = new Row[docRows.Count];

                                for (var j = 0; j < docRows.Count - 1; j++)
                                {
                                    var cells = new Cell[docRows[j].ChildNodes.Count];

                                    for (var i = 0; i < cells.Count(); i++)
                                    {
                                        var location = docRows[j].ChildNodes[i].Attributes.GetNamedItem("r").Value;
                                        var value = docRows[j].ChildNodes[i].ChildNodes[0].InnerText;
                                        var cell = new Cell(location, value);
                                        cells[i] = cell;
                                    }

                                    var row = new Row(j + 1,cells);

                                    rows[j] = row;
                                }

                                var worksheet = new Worksheet(docWorksheets[worksheetCounter],  rows);
                                worksheetList.Add(worksheet);
                                worksheetCounter++;
                            }
                           
                        }
                       
                    }
                    this.Worksheets = worksheetList.ToArray();
                }
            } else
            {
               doc.LoadXml(Encoding.UTF8.GetString(this.File));
            }

            



        }

        public Workbook(Byte[] file)
        {
            if (File != null)
            {
                this.File = file;
            } else
            {
                throw new FileLoadException("Given file is empty");
            }
        }

        public Workbook(String filename)
        {
            if (filename == null)
            {
                throw new FileNotFoundException("Filename is empty");
            } else
            {
                if (!System.IO.File.Exists(filename))
                    throw new FileNotFoundException("File '" + filename + "' wasn't found.");
            }

            this.Filename = filename;
                    
        }


    }

    


}
