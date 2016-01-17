# Net-OpenXml.Excel.Data
Small library for working with excel as with IDataReader. 
With this library you can easy read excel file to DataTable or you can import excel file to database (use SqlBulkCopy).

Read to DataTable example:

    class Program
    {
        static void Main(string[] args)
        {
            var reader = new ExcelDataReader(@"data.xlsx");
            var dt = new DataTable();
            
            dt.Load(reader);
            Console.WriteLine("done: " + dt.Rows.Count);
            Console.ReadKey();
        }
    }
