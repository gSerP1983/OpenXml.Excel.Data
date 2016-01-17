using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXml.Excel.Data.Util;

namespace UnitTest.OpenXml.Excel.Data
{
    [TestClass]
    public class ExcelUtilTest
    {
        [TestMethod]
        public void GetColumnIndexByNameTest()
        {
            Assert.AreEqual(ExcelUtil.GetColumnIndexByName("A1"), 0);
            Assert.AreEqual(ExcelUtil.GetColumnIndexByName("B32"), 1);
            Assert.AreEqual(ExcelUtil.GetColumnIndexByName("Z67Y"), 25);

            Assert.AreEqual(ExcelUtil.GetColumnIndexByName("AA1"), 26);
            Assert.AreEqual(ExcelUtil.GetColumnIndexByName("AE1T"), 30);
        }
    }
}
