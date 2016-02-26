using System;
using System.Linq;
using System.Data;
using System.Data.SqlClient;

namespace OpenXml.Excel.Data.Test
{
    class Program
    {
        private const string ConnectionString = @"Server=(local)\sqlexpress;Database=TestDb;Trusted_Connection=True;";
        private const string TableName = "ImportTable";

        static void Main()
        {
            DataTableBulkCopySample();

            // The best way to copy data to sql server
            DataReaderBulkCopySample();

            Console.ReadKey();
        }

        private static void DataTableBulkCopySample()
        {
            var dt = new DataTable();
            using (var reader = new ExcelDataReader(@"test.xlsx"))
                dt.Load(reader);
            Console.WriteLine("Read DataTable done: " + dt.Rows.Count);

            DataHelper.CreateTableIfNotExists(ConnectionString, TableName, dt.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToArray());
            Console.WriteLine("Create table done.");

            using (var bulkCopy = new SqlBulkCopy(ConnectionString))
            {
                bulkCopy.DestinationTableName = TableName;
                foreach (DataColumn dc in dt.Columns)
                    bulkCopy.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);

                bulkCopy.WriteToServer(dt);
            }
            Console.WriteLine("Copy data to database done (DataTable).");
        }

        private static void DataReaderBulkCopySample()
        {            
            using (var reader = new ExcelDataReader(@"test.xlsx"))
            {
                var cols = Enumerable.Range(0, reader.FieldCount).Select(i => reader.GetName(i)).ToArray();

                DataHelper.CreateTableIfNotExists(ConnectionString, TableName, cols);
                Console.WriteLine("Create table done.");

                using (var bulkCopy = new SqlBulkCopy(ConnectionString))
                {
                    // MSDN: When EnableStreaming is true, SqlBulkCopy reads from an IDataReader object using SequentialAccess, 
                    // optimizing memory usage by using the IDataReader streaming capabilities
                    bulkCopy.EnableStreaming = true;

                    bulkCopy.DestinationTableName = TableName;
                    foreach (var col in cols)
                        bulkCopy.ColumnMappings.Add(col, col);

                    bulkCopy.WriteToServer(reader);
                }
                Console.WriteLine("Copy data to database done (DataReader).");
            }
        }
    }
}