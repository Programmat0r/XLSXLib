namespace XLSXLib
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Compression;
    using System.Linq;
    using System.Runtime.CompilerServices;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// Defines the <see cref="Workbook" />.
    /// </summary>
    public class Workbook
    {
        /// <summary>
        /// Gets or sets the File.
        /// </summary>
        public Byte[] File { get; set; }

        /// <summary>
        /// Gets or sets the Filename.
        /// </summary>
        public String Filename { get; set; }

        /// <summary>
        /// Gets the Name.
        /// </summary>
        public String Name { get; }

        /// <summary>
        /// Gets the Worksheets.
        /// </summary>
        public Worksheet[] Worksheets { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Workbook"/> class.
        /// </summary>
        public Workbook()
        {
        }


        /// <summary>
        /// The Load.
        /// </summary>
        public void Load()
        {
            var doc = new XmlDocument();

            if (this.Filename != null)
            {

                using (var data = System.IO.File.OpenRead(this.Filename))

                using (var zip = new ZipArchive(data, ZipArchiveMode.Read))
                {
                    var docWorkbook = zip.GetEntry("xl/workbook.xml");

                    var docWorksheets = new List<String>();

                    using (var stream = docWorkbook.Open())
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
                        if (entry.FullName.Contains("xl/worksheets/"))
                        {

                            using (var stream = entry.Open())
                            {
                                var sr = new StreamReader(stream);
                                doc.LoadXml(sr.ReadToEnd());

                                var docRows = doc.GetElementsByTagName("row");

                                var rows = new Row[docRows.Count];

                                for (var j = 0; j < docRows.Count; j++)
                                {
                                    var cells = new Cell[docRows[j].ChildNodes.Count];

                                    for (var i = 0; i < cells.Count(); i++)
                                    {
                                        var location = docRows[j].ChildNodes[i].Attributes.GetNamedItem("r").Value;
                                        var value = "";
                                        if (docRows[j].ChildNodes[i].HasChildNodes)
                                             value = docRows[j].ChildNodes[i].ChildNodes[0].InnerText;
                                        cells[i] = new Cell(location, value);
                                       
                                    }

                                    var row = new Row(j + 1, cells);

                                    rows[j] = row;
                                }

                                var worksheet = new Worksheet(entry.FullName, docWorksheets[worksheetCounter], rows);
                            
                                worksheetList.Add(worksheet);
                                worksheetCounter++;
                            }

                        }

                    }
                    this.Worksheets = worksheetList.ToArray();
                }
            }
            else
            {
                doc.LoadXml(Encoding.UTF8.GetString(this.File));
            }
        }

        public void Save()
        {
            
            if(this.Filename != null)
            {

                using (var data = System.IO.File.Open(this.Filename, FileMode.Open))

                using (var zip = new ZipArchive(data, ZipArchiveMode.Update))
                {


                    foreach (var entry in zip.Entries)
                    {
                        if (entry.FullName.Contains("xl/worksheets/"))
                        {

                             
                                var sr = new StreamReader(entry.Open());
                                var doc = new XmlDocument();
                                var xmlString = sr.ReadToEnd();
                                sr.Close();

                                var stream = entry.Open();
                                doc.LoadXml(xmlString);


                            bool changes = false;

                            foreach (var worksheet in this.Worksheets)
                                {
                                
                                    if (entry.FullName == worksheet.SystemName)
                                    {
                                        foreach (var row in worksheet.Rows)
                                        {
                                            var docCells = doc.GetElementsByTagName("c");
                                            foreach (XmlNode dR in docCells)
                                            {
                                                foreach (var cell in row.Cells)
                                                {
                                                    if (dR.Attributes.GetNamedItem("r").Value == cell.Location)
                                                    {
                                                        if (dR.HasChildNodes)
                                                        {
                                                            if(dR.ChildNodes[0].InnerText != cell.Value)
                                                        {
                                                            dR.ChildNodes[0].InnerText = cell.Value;
                                                            changes = true;
                                                        }
                                                            
                                                        }
                                                        else
                                                        {
                                                            var v = doc.CreateElement("v");
                                                            v.InnerText = cell.Value;
                                                            dR.AppendChild(v);
                                                            changes = true;
                                                        }
                                                        
                                                    }
                                                }
                                            }

                                        }
                                    }
                                }
                                //var sW = new StreamWriter(stream);
                                //sW.BaseStream.
                                if (changes)
                                    doc.Save(stream);
                                //doc.Save(@"C:\Users\theni\Desktop\test.xml");


                            
                        }
                    }
                }

             }

        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Workbook"/> class.
        /// </summary>
        /// <param name="file">The file<see cref="Byte[]"/>.</param>
        public Workbook(Byte[] file)
        {
            if (File != null)
            {
                this.File = file;
            }
            else
            {
                throw new FileLoadException("Given file is empty");
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Workbook"/> class.
        /// </summary>
        /// <param name="filename">The filename<see cref="String"/>.</param>
        public Workbook(String filename)
        {
            if (filename == null)
            {
                throw new FileNotFoundException("Filename is empty");
            }
            else
            {
                if (!System.IO.File.Exists(filename))
                    throw new FileNotFoundException("File '" + filename + "' wasn't found.");
            }

            this.Filename = filename;
        }
    }
}
