using HtmlAgilityPack;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Text;
using TowerSoft.HtmlToExcel;

namespace HtmlToExcelTests {
    [TestClass]
    public class HtmlAgilityUtilitiesTest {
        [TestMethod]
        public void GetHtmlTableNode_HtmlWithSingleTable_ShouldReturnTableHtmlNode() {
            string htmlWithTable = "<div><table><thead></thead><tbody></tbody></table></div>";
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(htmlWithTable);

            HtmlNode node = new HtmlAgilityUtilities().GetHtmlTableNode(htmlDoc);
            Assert.IsTrue(node != null && node.Name == "table");

            if (node == null)
                Assert.Fail("The HtmlNode was null");
            if (node.Name != "table")
                Assert.Fail("The HtmlNode was not a table");
        }

        [TestMethod]
        public void GetHtmlTableNode_HtmlWithMultipleTables_ShouldThrowException() {
            string htmlWithTable = "<table></table><table></table>";
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(htmlWithTable);

            try {
                HtmlNode node = new HtmlAgilityUtilities().GetHtmlTableNode(htmlDoc);
            } catch (Exception ex) {
                StringAssert.Contains(ex.Message, HtmlAgilityUtilities.MultipleTableNodesFoundMessage);
            }
        }

        [TestMethod]
        public void GetHtmlTableNode_HtmlWithNoTables_ShouldThrowException() {
            string htmlWithTable = "<div></div>";
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(htmlWithTable);

            try {
                HtmlNode node = new HtmlAgilityUtilities().GetHtmlTableNode(htmlDoc);
            } catch (Exception ex) {
                StringAssert.Contains(ex.Message, HtmlAgilityUtilities.NoTableNodesFoundMessage);
            }
        }
    }
}
