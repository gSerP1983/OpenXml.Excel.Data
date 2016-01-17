using System;
using System.Data;

namespace OpenXml.Excel.Data.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            var dt = new DataTable();
            using (var reader = new ExcelDataReader(@"C:\Users\s-petrov.COMPULINK\Desktop\excelImport\xlsx\05 НУК 1501-3000-10000.xlsx"))
            {                
                dt.Load(reader);
            }

            Console.WriteLine("done: " + dt.Rows.Count);
            Console.ReadKey();
        }
    }
}
