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

            WorkbookBuilder workbookBuilder = new();
            workbookBuilder.AddSheet("test", html);
            workbookBuilder.AddSheet("test", html);
            workbookBuilder.AddSheet("test", html);
            workbookBuilder.AddSheet("test2 ", html);
            workbookBuilder.AddSheet("test2", html);
            workbookBuilder.AddSheet("Loremipsumdolorsitam_consecteturadipiscin", html);
            workbookBuilder.AddSheet("Loremipsumdolorsitam_consecteturadipiscin", html);

            workbookBuilder.AddSheet("illegal-characters-/\\*?[]", html);
            workbookBuilder.AddSheet("/\\*?[]", html);

            workbookBuilder.AddSheet("sheet2", html);

            byte[] data = workbookBuilder.GetAsByteArray();

            File.WriteAllBytes(Path.Combine(Environment.CurrentDirectory, "builderTest.xlsx"), data);
        }
    }
}
