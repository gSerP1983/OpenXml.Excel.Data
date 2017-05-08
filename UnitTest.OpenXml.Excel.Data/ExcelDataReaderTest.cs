using System;
using System.Data;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXml.Excel.Data;

namespace UnitTest.OpenXml.Excel.Data
{
    [TestClass]
    public class ExcelDataReaderTest
    {
        [TestMethod]
        public void OpenPathEmptyRowsTest()
        {
            var dt = new DataTable();
            using (var reader = new ExcelDataReader(@"NoDataReturned.xlsx", "ContractIBNR"))
                dt.Load(reader);

            Assert.AreEqual(5, dt.Columns.Count);
            Assert.AreEqual(1, dt.Rows.Count);
        }

        [TestMethod]
        public void OpenPathHdrDefaultSheetTest()
        {
            var dt = new DataTable();
            using (var reader = new ExcelDataReader(@"test.xlsx"))
                dt.Load(reader);

            Assert.AreEqual(7, dt.Columns.Count);
            Assert.AreEqual(5, dt.Rows.Count);
        }

        [TestMethod]
        public void OpenPathNoHdrSheetNameTest()
        {
            var dt = new DataTable();
            using (var reader = new ExcelDataReader(@"test.xlsx", "Second", false))
                dt.Load(reader);

            Assert.AreEqual(7, dt.Columns.Count);
            Assert.AreEqual(6, dt.Rows.Count);
        }

        [TestMethod]
        public void OpenStreamHdrSheetIndexTest()
        {
            var dt = new DataTable();

            using (var sr = File.OpenRead(@"test.xlsx"))
            {
                using (var reader = new ExcelDataReader(sr, 1, false))
                    dt.Load(reader);
            }

            Assert.AreEqual(7, dt.Columns.Count);
            Assert.AreEqual(6, dt.Rows.Count);
        }

        [TestMethod]
        public void OpenStreamEmptySheetTest()
        {
            var dt = new DataTable();

            using (var sr = File.OpenRead(@"test.xlsx"))
            {
                using (var reader = new ExcelDataReader(sr, "Third", false))
                    dt.Load(reader);
            }

            Assert.AreEqual(0, dt.Columns.Count);
            Assert.AreEqual(0, dt.Rows.Count);
        }

        [TestMethod]
        [ExpectedException(typeof(ApplicationException))]
        public void OpenInvalidSheetIndexTest()
        {
            using (new ExcelDataReader(@"test.xlsx", 27))
            {
            }
        }

        [TestMethod]
        [ExpectedException(typeof(ApplicationException))]
        public void OpenInvalidSheetNameTest()
        {
            using (new ExcelDataReader(@"test.xlsx", "Yuri Gagarin"))
            {
            }
        }

        [TestMethod]
        public void DataReaderTest()
        {
            using (var reader = new ExcelDataReader(@"test.xlsx"))
            {
                Assert.AreEqual(0, reader.Depth);
                Assert.AreEqual(-1, reader.RecordsAffected);
                Assert.AreEqual(7, reader.FieldCount);
                Assert.AreEqual(5, reader.GetOrdinal("ColDecimal"));
                Assert.AreEqual(-1, reader.GetOrdinal("InvalidColName"));
            }
        }

        [TestMethod]
        public void ReadTest()
        {
            using (var reader = new ExcelDataReader(@"test.xlsx"))
            {
                reader.Read();
                Assert.AreEqual(381728, reader.GetInt32(0));
                Assert.AreEqual(381728, reader.GetInt64(0));

                Assert.AreEqual("Mr Brown", reader.GetString(1));
                Assert.AreEqual("Mr Brown", reader[1]);
                Assert.AreEqual("Mr Brown", reader["Name"]);

                Assert.AreEqual(new DateTime(1983, 3, 27, 6, 55, 0), reader.GetDateTime(2));

                Assert.AreEqual(new Guid("6E2BF784-F116-494A-916D-9DFF9B2A2AA0"), reader.GetGuid(3));

                Assert.AreEqual(32, reader.GetInt16(4));
                Assert.AreEqual(32, reader.GetInt32(4));
                Assert.AreEqual(32, reader.GetInt64(4));
                Assert.AreEqual(32, reader.GetByte(4));

                Assert.AreEqual(917.68m, reader.GetDecimal(5));
                Assert.AreEqual(917.68d, reader.GetDouble(5));
                Assert.AreEqual(917.68f, reader.GetFloat(5));

                Assert.AreEqual(true, reader.GetBoolean(6));
            }
        }
    }
}
