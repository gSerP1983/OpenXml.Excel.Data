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
        public void OpenPathHdrDefaultSheet()
        {
            var dt = new DataTable();
            using (var reader = new ExcelDataReader(@"test.xlsx"))
                dt.Load(reader);

            Assert.AreEqual(7, dt.Columns.Count);
            Assert.AreEqual(5, dt.Rows.Count);
        }

        [TestMethod]
        public void OpenPathNoHdrSheetName()
        {
            var dt = new DataTable();
            using (var reader = new ExcelDataReader(@"test.xlsx", "Second", false))
                dt.Load(reader);

            Assert.AreEqual(7, dt.Columns.Count);
            Assert.AreEqual(6, dt.Rows.Count);
        }

        [TestMethod]
        public void OpenStreamHdrSheetIndex()
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
        public void OpenStreamEmptySheet()
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
        public void OpenInvalidSheetIndex()
        {
            using (new ExcelDataReader(@"test.xlsx", 27))
            {
            }
        }

        [TestMethod]
        [ExpectedException(typeof(ApplicationException))]
        public void OpenInvalidSheetName()
        {
            using (new ExcelDataReader(@"test.xlsx", "Yuri Gagarin"))
            {
            }
        }
    }
}
