using HtmlAgilityPack;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace TowerSoft.HtmlToExcel {
    internal class EPPlusUtilities {
        private HtmlToExcelSettings Settings { get; }

        private bool _hasMergedCells = false;

        internal EPPlusUtilities(HtmlToExcelSettings settings) {
            Settings = settings;
        }

        internal byte[] GenerateWorkbookFromHtmlNode(HtmlNode node) {
            using (ExcelPackage package = new ExcelPackage()) {
                CreateSheet(package, "Sheet", node);
                return package.GetAsByteArray();
            }
        }

        internal void CreateSheet(ExcelPackage package, string sheetName, HtmlNode node) {
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Sheet");

            int row = 1;
            int col = 1;
            foreach (HtmlNode rowNode in node.Descendants().Where(x => x.Name == "tr")) {

                List<HtmlNode> cells = rowNode.Elements("td").ToList();
                cells.AddRange(rowNode.Elements("th"));
                foreach (HtmlNode cellNode in cells) {
                    RenderCell(sheet, cellNode, ref row, ref col);
                }
                col = 1;
                row++;
            }

            if (!_hasMergedCells) {
                var table = sheet.Tables.Add(sheet.Cells[sheet.Dimension.Address], "mainTable" + sheet.Index);
                table.TableStyle = OfficeOpenXml.Table.TableStyles.Light1;
                table.ShowRowStripes = Settings.ShowRowStripes;
                table.ShowFilter = Settings.ShowFilter;
            }

            if (Settings.AutofitColumns) {
                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
            }
        }

        private void RenderCell(ExcelWorksheet sheet, HtmlNode cellNode, ref int row, ref int col) {
            ExcelRange cell = sheet.Cells[row, col];
            cell.Value = cellNode.InnerText.SafeTrim();

            if (cellNode.Name == "th") { // Set font bold for th elements
                cell.Style.Font.Bold = true;
            }

            if (cellNode.HasAttributes) {
                HtmlAttribute boldAttribute = cellNode.Attributes.SingleOrDefault(x => x.Name == "data-excel-bold");
                if (boldAttribute != null) {
                    if (bool.TryParse(boldAttribute.Value, out bool isBold)) {
                        cell.Style.Font.Bold = isBold;
                    }
                }

                HtmlAttribute hyperlinkAttribute = cellNode.Attributes.SingleOrDefault(x => x.Name == "data-excel-hyperlink");
                if (hyperlinkAttribute != null) {
                    if (Uri.TryCreate(hyperlinkAttribute.Value, UriKind.Absolute, out Uri uri)) {
                        cell.Hyperlink = uri;
                        cell.Style.Font.Color.SetColor(Color.Blue);
                        cell.Style.Font.UnderLine = true;
                    } else {
                        cell.AddComment("Unable to parse hyperlink: " + hyperlinkAttribute.Value, "TowerSoft.HtmlToExcel");
                    }
                }

                HtmlAttribute commentAttribute = cellNode.Attributes.SingleOrDefault(x => x.Name == "data-excel-comment");
                if (commentAttribute != null && !string.IsNullOrWhiteSpace(commentAttribute.Value)) {
                    string author = "System";
                    HtmlAttribute authorAttribute = cellNode.Attributes.SingleOrDefault(x => x.Name == "data-excel-comment-author");
                    if (authorAttribute != null && !string.IsNullOrWhiteSpace(authorAttribute.Value)) {
                        author = authorAttribute.Value;
                    }
                    cell.AddComment(commentAttribute.Value, author);
                }
            }

            int colspan = cellNode.GetAttributeValue("colspan", 1);
            if (colspan > 1) {
                sheet.Cells[row, col, row, col + colspan - 1].Merge = true;
                _hasMergedCells = true;
                col += colspan;
            } else {
                col++;
            }
        }
    }
}
