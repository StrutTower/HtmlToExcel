using AngleSharp.Dom;
using System;
using System.Linq;

namespace TowerSoft.HtmlToExcel {
    internal class HtmlAgilityUtilities {
        internal const string NoTableNodesFoundMessage = "The supplied HtmlDocument did not have a table element.";
        internal const string MultipleTableNodesFoundMessage = "The supplied HtmlDocument has more than one table element.";

        internal IElement GetHtmlTableElement(IElement htmlDoc) {
            var tables = htmlDoc.QuerySelectorAll("table");
            if (tables.Count() < 1) {
                throw new Exception(NoTableNodesFoundMessage);
            }
            if (tables.Count() > 1) {
                throw new Exception(MultipleTableNodesFoundMessage);
            }
            return tables.First();
        }
    }
}
