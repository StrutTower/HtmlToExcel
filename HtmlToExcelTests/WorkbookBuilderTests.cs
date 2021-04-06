using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using TowerSoft.HtmlToExcel;

namespace HtmlToExcelTests {
    [TestClass]
    public class WorkbookBuilderTests {
        [TestMethod]
        public void OutputTestWorkbook() {
            string html = File.ReadAllText("htmlTable.html");

            WorkbookBuilder workbookBuilder = new WorkbookBuilder();
            workbookBuilder.AddSheet("test", html);

            workbookBuilder.AddSheet("sheet2", html);

            byte[] data = workbookBuilder.GetAsByteArray();

            File.WriteAllBytes(Path.Combine(Environment.CurrentDirectory, "buiilderTest.xlsx"), data);
        }
    }
}
