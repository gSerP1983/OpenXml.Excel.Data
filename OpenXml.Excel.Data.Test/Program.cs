using System;
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
            DataTableSample();
            Console.ReadKey();
        }

        private static void DataTableSample()
        {
            // read to Datatable sample
            var dt = new DataTable();
            using (var reader = new ExcelDataReader(@"test.xlsx"))
                dt.Load(reader);
            Console.WriteLine("Read DataTable done: " + dt.Rows.Count);

            // create table in database by DataTable
            DataHelper.CreateTableIfNotExists(ConnectionString, TableName, dt);
            Console.WriteLine("Create database table done.");

            // very fast way to copy data to sql server
            using (var bulkCopy = new SqlBulkCopy(ConnectionString))
            {
                bulkCopy.DestinationTableName = TableName;
                foreach (DataColumn dc in dt.Columns)
                    bulkCopy.ColumnMappings.Add(dc.ColumnName, dc.ColumnName);

                bulkCopy.WriteToServer(dt);
            }
            Console.WriteLine("Copy data ti database done.");
        }
    }
}