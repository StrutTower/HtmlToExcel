using TowerSoft.HtmlToExcel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace HtmlToExcelTests {
    [TestClass]
    public class WorkbookGeneratorTests {
        [TestMethod]
        public void TestWorkbook() {
            string html = File.ReadAllText("htmlTable.html");

            var data = new WorkbookGenerator().FromHtmlString(html);

            File.WriteAllBytes(Path.Combine(Environment.CurrentDirectory, "test.xlsx"), data);
        }
    }
}