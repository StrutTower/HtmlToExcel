using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;

namespace TowerSoft.HtmlToExcel {
    internal class HtmlAgilityUtilities {
        internal const string NoTableNodesFoundMessage = "The supplied HtmlDocument did not have a table element.";
        internal const string MultipleTableNodesFoundMessage = "The supplied HtmlDocument has more than one table element.";

        internal HtmlNode GetHtmlTableNode(HtmlDocument html) {
            List<HtmlNode> nodes = html.DocumentNode.Descendants("table").ToList();
            if (nodes.Count < 1) {
                throw new Exception(NoTableNodesFoundMessage);
            }
            if (nodes.Count > 1) {
                throw new Exception(MultipleTableNodesFoundMessage);
            }
            return nodes.First();
        }
    }
}
