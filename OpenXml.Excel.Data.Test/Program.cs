using System;
using System.Data;

namespace OpenXml.Excel.Data.Test
{
    class Program
    {
        static void Main()
        {
            var dt = new DataTable();
            using (var reader = new ExcelDataReader(@"test.xlsx"))           
                dt.Load(reader);

            Console.WriteLine("done: " + dt.Rows.Count);
            Console.ReadKey();
        }
    }
}
