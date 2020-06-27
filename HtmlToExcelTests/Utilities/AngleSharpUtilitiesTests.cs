using AngleSharp;
using AngleSharp.Dom;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Text;
using TowerSoft.HtmlToExcel.Utilities;

namespace HtmlToExcelTests.Utilities {
    [TestClass]
    public class AngleSharpUtilitiesTests {
        [TestMethod]
        public void GetHtmlTableNode_HtmlWithSingleTable_ShouldReturnTableHtmlNode() {
            string htmlWithTable = "<div><table><thead></thead><tbody></tbody></table></div>";
            IBrowsingContext ctx = BrowsingContext.New(Configuration.Default.WithDefaultLoader());
            var doc = ctx.OpenAsync(r => r.Content(htmlWithTable)).Result;

            IElement element = new AngleSharpUtilities().GetHtmlTableNode(doc.DocumentElement);

            Assert.IsTrue(element != null && element.NodeName.Equals("table", StringComparison.InvariantCultureIgnoreCase));
        }

        [TestMethod]
        public void GetHtmlTableNode_HtmlWithMultipleTables_ShouldThrowException() {
            string htmlWithTable = "<table></table><table></table>";
            IBrowsingContext ctx = BrowsingContext.New(Configuration.Default.WithDefaultLoader());
            var doc = ctx.OpenAsync(r => r.Content(htmlWithTable)).Result;

            try {
                IElement element = new AngleSharpUtilities().GetHtmlTableNode(doc.DocumentElement);
            } catch (Exception ex) {
                StringAssert.Contains(ex.Message, AngleSharpUtilities.MultipleTableNodesFoundMessage);
            }
        }

        [TestMethod]
        public void GetHtmlTableNode_HtmlWithNoTables_ShouldThrowException() {
            string htmlWithTable = "<div></div>";
            IBrowsingContext ctx = BrowsingContext.New(Configuration.Default.WithDefaultLoader());
            var doc = ctx.OpenAsync(r => r.Content(htmlWithTable)).Result;

            try {
                IElement element = new AngleSharpUtilities().GetHtmlTableNode(doc.DocumentElement);
            } catch (Exception ex) {
                StringAssert.Contains(ex.Message, AngleSharpUtilities.NoTableNodesFoundMessage);
            }
        }
    }
}
