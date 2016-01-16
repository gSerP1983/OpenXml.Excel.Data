using System;
using System.Data;

namespace OpenXml.Excel.Data.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            var reader = new ExcelDataReader(@"C:\Users\s-petrov.COMPULINK\Desktop\excelImport\xlsx\05 НУК 1501-3000.xlsx");
            var dt = new DataTable();

            dt.Load(reader);
            Console.WriteLine("done: " + dt.Rows.Count);
            Console.ReadKey();
        }
    }
}
