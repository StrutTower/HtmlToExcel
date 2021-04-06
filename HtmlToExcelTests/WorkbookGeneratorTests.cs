using TowerSoft.HtmlToExcel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace HtmlToExcelTests {
    [TestClass]
    public class WorkbookGeneratorTests {
        [TestMethod]
        public void OutputTestWorkbook() {
            string html = File.ReadAllText("htmlTable.html");

            byte[] data = new WorkbookGenerator().FromHtmlString(html);

            File.WriteAllBytes(Path.Combine(Environment.CurrentDirectory, "test.xlsx"), data);
        }
    }
}