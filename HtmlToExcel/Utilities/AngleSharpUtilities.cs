﻿using AngleSharp;
using AngleSharp.Dom;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TowerSoft.HtmlToExcel.Utilities {
    internal class AngleSharpUtilities {
        internal const string NoTableNodesFoundMessage = "The supplied HtmlDocument did not have a table element.";
        internal const string MultipleTableNodesFoundMessage = "The supplied HtmlDocument has more than one table element.";

        internal IElement GetHtmlTableNode(IElement html) {
            List<IElement> nodes = html.QuerySelectorAll("table").ToList();
            if (nodes.Count() < 1) {
                throw new Exception(NoTableNodesFoundMessage);
            }
            if (nodes.Count > 1) {
                throw new Exception(MultipleTableNodesFoundMessage);
            }
            return nodes.First();
        }
    }
}
