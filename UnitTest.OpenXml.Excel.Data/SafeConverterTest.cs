using System.Globalization;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenXml.Excel.Data.Util;

namespace UnitTest.OpenXml.Excel.Data
{
    [TestClass]
    public class SafeConverterTest
    {

        [TestMethod]
        public void ConvertTest()
        {
            var value = SafeConverter.Convert(0, typeof (bool));
            Assert.AreEqual(value.GetType(), typeof(bool));
            Assert.AreEqual(value.ToString(), "False");

            Assert.AreEqual(SafeConverter.Convert<bool>(0), false);
            Assert.AreEqual(SafeConverter.Convert<bool>("0"), false);
            Assert.AreEqual(SafeConverter.Convert<bool>("1"), true);

            Assert.AreEqual(SafeConverter.Convert<int>("1"), 1);

            var separator = CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator;
            var stringDecimal = "1,89".Replace(".", separator).Replace(",", separator);
            Assert.AreEqual(SafeConverter.Convert<double>(stringDecimal), 1.89d);

            value = SafeConverter.Convert(0, typeof(string));
            Assert.AreEqual(value.GetType(), typeof(string));

            Assert.AreEqual(SafeConverter.Convert<int?>("1"), 1);
            Assert.AreEqual(SafeConverter.Convert<int?>(null), null);
            Assert.AreEqual(SafeConverter.Convert<int?>(""), null);
            Assert.AreEqual(SafeConverter.Convert(null, typeof(int?)), null);

            Assert.AreEqual(SafeConverter.Convert<int>(BindingFlags.Static), 8);
            Assert.AreEqual(SafeConverter.Convert<BindingFlags>(8), BindingFlags.Static);
            Assert.AreEqual(SafeConverter.Convert<BindingFlags>("8"), BindingFlags.Static);
        }
    }
}
