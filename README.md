# OpenXml.Excel.Data
A simple C# IDataReader implementation to read Excel Open Xml files.
ExcelDataReader reads very fast. On my machine 10000 records per 3 sec.
ExcelDataReader uses small memory because it reads SAX method (OpenXmlReader).
With this library you can easy read excel file to DataTable or you can import excel file to sql server database (use SqlBulkCopy).
Enjoy))

Read to DataTable example:

    class Program
    {
        static void Main(string[] args)
        {
            var dt = new DataTable();
            using (var reader = new ExcelDataReader(@"data.xlsx"))
            {                
                dt.Load(reader);
            }

            Console.WriteLine("done: " + dt.Rows.Count);
            Console.ReadKey();
        }
    }