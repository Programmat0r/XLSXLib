# XLSXLib
A lightweight .NET Library to read XLSX Files.

<p>Not finished. Do not use this library in production!</p>

<h2>Requirements</h2>
<p><b>.Net Framework 4.7.2</b></p>
<p>This library was programmed in .Net Framework 4.7.2. You can download it and decrease the .Net version as much as you need it. I didn't tested it below the mentioned version and don't know if it will work.</p>

<h2>Example</h2>
<p>Read all cells of the first worksheet of a workbook</p>

```csharp
var workbook = new Workbook("Example/Testfile.xlsx");
workbook.Load();

foreach (var row in workbook.Worksheets[0].Rows)
{
    foreach (var cell in row.Cells)
    {
        Console.Write(cell.Value;

    }
    Console.WriteLine("");
}
```

<h2>License</h2>

<p>Code - <a href="http://www.apache.org/licenses/LICENSE-2.0">APACHE LICENSE, VERSION 2.0</a></p>
