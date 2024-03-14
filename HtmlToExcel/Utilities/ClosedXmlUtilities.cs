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
            bool valueSet = false;
            string value = cellNode.TextContent.SafeTrim();

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
                
                IAttr horizontalAlignmentAttribute = cellNode.Attributes.SingleOrDefault(x => x.Name == "horizontal-alignment");
                if (horizontalAlignmentAttribute != null) {
                    if (Enum.TryParse(horizontalAlignmentAttribute.Value, true, out XLAlignmentHorizontalValues horizontalAlignment))
                    {
                        cell.Style.Alignment.Horizontal = horizontalAlignment;
                    }
                }

                IAttr hyperlinkAttribute = cellNode.Attributes.SingleOrDefault(x => x.Name == "data-excel-hyperlink");
                if (hyperlinkAttribute != null) {
                    if (Uri.TryCreate(hyperlinkAttribute.Value, UriKind.Absolute, out Uri uri)) {
                        cell.SetHyperlink(new XLHyperlink(uri));
                    } else {
                        cell.CreateComment()
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
                        cell.CreateComment().SetAuthor(author).AddSignature();
                    }
                    cell.CreateComment().AddText(commentAttribute.Value);
                }

                IAttr dataTypeAttribute = cellNode.Attributes.SingleOrDefault(x => x.Name == "data-type");
                IAttr dataFormatAttribute = cellNode.Attributes.SingleOrDefault(x => x.Name == "data-format");
                if (dataTypeAttribute != null && !string.IsNullOrWhiteSpace(dataTypeAttribute.Value)) {
                    switch (dataTypeAttribute.Value.ToLower()) {
                        case "text":
                        case "string":
                            cell.SetValue(value);
                            valueSet = true;
                            break;
                        case "number":
                        case "int":
                        case "double":
                        case "float":
                        case "decimal":
                            if (decimal.TryParse(value, out decimal number)) {
                                if (dataFormatAttribute != null && !string.IsNullOrWhiteSpace(dataFormatAttribute.Value))
                                    cell.Style.NumberFormat.Format = dataFormatAttribute.Value;
                                cell.SetValue(number);
                                valueSet = true;
                            }
                            break;
                        case "bool":
                        case "boolean":
                            if (bool.TryParse(value, out bool boolValue)) {
                                cell.SetValue(boolValue);
                                valueSet = true;
                            }
                            break;
                        case "date":
                        case "datetime":
                            if (DateTime.TryParse(value, out DateTime dateTime)) {
                                if (dataFormatAttribute != null && !string.IsNullOrWhiteSpace(dataFormatAttribute.Value))
                                    cell.Style.DateFormat.Format = dataFormatAttribute.Value;
                                cell.SetValue(dateTime);
                                valueSet = true;
                            }
                            break;
                        case "time":
                        case "timespan":
                            if (TimeSpan.TryParse(value, out TimeSpan timeSpan)) {
                                cell.SetValue(timeSpan);
                                valueSet = true;
                            }
                            break;
                    }
                }
            }

            if (!valueSet) {
                cell.SetValue(value);
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
            
            if (int.TryParse(cellNode.GetAttribute("font-size"), out int fontSize) && fontSize > 0) {
                cell.Style.Font.FontSize = fontSize;
            }
        }
    }
}
