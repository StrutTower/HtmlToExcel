using AngleSharp.Dom;
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

        internal byte[] GenerateWorkbookFromHtmlNode(IElement tableNode) {
            ExcelPackage.LicenseContext = Settings.EpplusLicenseContext;
            using (ExcelPackage package = new ExcelPackage()) {
                CreateSheet(package, "Sheet", tableNode);
                return package.GetAsByteArray();
            }
        }

        internal void CreateSheet(ExcelPackage package, string sheetName, IElement tableNode) {
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add(sheetName);

            int row = 1;
            int col = 1;
            foreach (IElement rowNode in tableNode.QuerySelectorAll("tr")) {

                List<IElement> cells = rowNode.QuerySelectorAll("th").ToList();
                cells.AddRange(rowNode.QuerySelectorAll("td"));
                foreach (IElement cellNode in cells) {
                    RenderCell(sheet, cellNode, ref row, ref col);
                }
                col = 1;
                row++;
            }

            if (!_hasMergedCells && sheet.Dimension != null) {
                var table = sheet.Tables.Add(sheet.Cells[sheet.Dimension.Address], "mainTable" + sheet.Index);
                table.TableStyle = OfficeOpenXml.Table.TableStyles.Light1;
                table.ShowRowStripes = Settings.ShowRowStripes;
                table.ShowFilter = Settings.ShowFilter;
            }

            if (Settings.AutofitColumns && sheet.Dimension != null) {
                sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
            }
        }

        private void RenderCell(ExcelWorksheet sheet, IElement cellNode, ref int row, ref int col) {
            ExcelRange cell = sheet.Cells[row, col];
            cell.Value = cellNode.ChildNodes.OfType<IText>().Select(m => m.Text).FirstOrDefault();

            if (cellNode.NodeName == "th") { // Set font bold for th elements
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
                        cell.Hyperlink = uri;
                        cell.Style.Font.Color.SetColor(Color.Blue);
                        cell.Style.Font.UnderLine = true;
                    } else {
                        cell.AddComment("Unable to parse hyperlink: " + hyperlinkAttribute.Value, "TowerSoft.HtmlToExcel");
                    }
                }

                IAttr commentAttribute = cellNode.Attributes.SingleOrDefault(x => x.Name == "data-excel-comment");
                if (commentAttribute != null && !string.IsNullOrWhiteSpace(commentAttribute.Value)) {
                    string author = "System";
                    IAttr authorAttribute = cellNode.Attributes.SingleOrDefault(x => x.Name == "data-excel-comment-author");
                    if (authorAttribute != null && !string.IsNullOrWhiteSpace(authorAttribute.Value)) {
                        author = authorAttribute.Value;
                    }
                    cell.AddComment(commentAttribute.Value, author);
                }
            }

            if (int.TryParse(cellNode.GetAttribute("colspan"), out int colspan)) {
                if (colspan > 1) {
                    sheet.Cells[row, col, row, col + colspan - 1].Merge = true;
                    _hasMergedCells = true;
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
