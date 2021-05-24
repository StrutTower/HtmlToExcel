using AngleSharp.Dom;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace TowerSoft.HtmlToExcel.Utilities {
    internal class ClosedXmlUtilities {
        private HtmlToExcelSettings Settings { get; }

        private bool hasMergedCells = false;

        internal ClosedXmlUtilities(HtmlToExcelSettings settings) {
            Settings = settings;
        }

        internal byte[] GenerateWorkbookFromHtmlNode(IElement tableNode) {
            using (IXLWorkbook workbook = new XLWorkbook()) {
                CreateWorksheet(workbook, "Sheet1", tableNode);

                using (MemoryStream stream = new MemoryStream()) {
                    workbook.SaveAs(stream);
                    return stream.ToArray();
                }
            }
        }

        internal void CreateWorksheet(IXLWorkbook workbook, string sheetName, IElement tableNode) {
            IXLWorksheet worksheet = workbook.Worksheets.Add(sheetName);

            int row = 1;
            int col = 1;
            foreach (IElement rowNode in tableNode.QuerySelectorAll("tr")) {
                List<IElement> cells = rowNode.QuerySelectorAll("th").ToList();
                cells.AddRange(rowNode.QuerySelectorAll("td"));
                foreach (IElement cellNode in cells) {
                    RenderCell(worksheet, cellNode, row, ref col);
                }
                col = 1;
                row++;
            }

            if (!hasMergedCells) {
                var table = worksheet.RangeUsed().CreateTable("mainTable" + worksheet.Name);
                table.Theme = XLTableTheme.TableStyleLight1;
                table.ShowRowStripes = Settings.ShowRowStripes;
                table.ShowAutoFilter = Settings.ShowFilter;
            }

            if (Settings.AutofitColumns) {
                worksheet.ColumnsUsed().AdjustToContents();
            }
        }

        private void RenderCell(IXLWorksheet worksheet, IElement cellNode, int row, ref int col) {
            IXLCell cell = worksheet.Cell(row, col);
            cell.Value = cellNode.TextContent.SafeTrim();

            if (cellNode.NodeName == "th") {
                cell.Style.Font.Bold = true;
            }

            if (cellNode.Attributes != null && cellNode.Attributes.Any()) {
                IAttr boldAttribute = cellNode.Attributes.SingleOrDefault(x => x.Name == "data-excel-bold");
                if (boldAttribute != null) {
                    if (bool.TryParse(boldAttribute.Value, out bool isBold)) {
                        cell.Style.Font.Bold = isBold;
                    }
                }

                IAttr hyperlinkAttribute = cellNode.Attributes.SingleOrDefault(x => x.Name == "data-excel-hyperlink");
                if (hyperlinkAttribute != null) {
                    if (Uri.TryCreate(hyperlinkAttribute.Value, UriKind.Absolute, out Uri uri)) {
                        cell.Hyperlink = new XLHyperlink(uri);
                    } else {
                        cell.Comment
                            .SetAuthor("TowerSoft.HtmlToExcel")
                            .AddSignature()
                            .AddText($"Unable to parse hyperlink: {hyperlinkAttribute.Value}");
                    }
                }

                IAttr commentAttribute = cellNode.Attributes.SingleOrDefault(x => x.Name == "data-excel-comment");
                if (commentAttribute != null && !string.IsNullOrWhiteSpace(commentAttribute.Value)) {
                    string author = "System";
                    IAttr authorAttribute = cellNode.Attributes.SingleOrDefault(x => x.Name == "data-excel-comment-author");
                    if (authorAttribute != null && !string.IsNullOrWhiteSpace(authorAttribute.Value)) {
                        author = authorAttribute.Value;
                        cell.Comment.SetAuthor(author).AddSignature();
                    }
                    cell.Comment.AddText(commentAttribute.Value);
                }
            }

            if (int.TryParse(cellNode.GetAttribute("colspan"), out int colspan)) {
                if (colspan > 1) {
                    worksheet.Range(worksheet.Cell(row, col), worksheet.Cell(row, col + colspan - 1)).Merge();
                    hasMergedCells = true;
                    col += colspan;
                } else {
                    col++;
                }
            } else {
                col++;
            }
        }
    }
}
