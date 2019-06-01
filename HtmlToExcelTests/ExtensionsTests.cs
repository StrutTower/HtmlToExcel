using Microsoft.VisualStudio.TestTools.UnitTesting;
using TowerSoft.HtmlToExcel;

namespace HtmlToExcelTests {
    [TestClass]
    public class ExtensionsTests {
        [TestMethod]
        public void SafeTrim_WithNullString_ShouldReturnEmptyString() {
            string nullString = null;

            string output = nullString.SafeTrim();

            Assert.AreEqual(string.Empty, output);
        }

        [TestMethod]
        public void SafeTrim_WithBlankString_ShouldReturnEmptyString() {
            string blankString = string.Empty;

            string output = blankString.SafeTrim();

            Assert.AreEqual(string.Empty, output);
        }

        [TestMethod]
        public void SafeTrim_WithTrimmableSpace_ShouldTrimString() {
            string trimmableString = "    test test    ";

            string result = trimmableString.SafeTrim();

            Assert.AreEqual("test test", result);
        }
    }
}
