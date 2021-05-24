using AngleSharp;
using AngleSharp.Dom;
using ClosedXML.Excel;
using System;
using System.IO;
using TowerSoft.HtmlToExcel.Utilities;

namespace TowerSoft.HtmlToExcel {
    /// <summary>
    /// Used to build a workbook with multiple sheets. Make sure to dispose of this class.
    /// </summary>
    public class WorkbookBuilder : IDisposable {
        private HtmlToExcelSettings Settings { get; }
        private IXLWorkbook Workbook { get; }

        /// <summary>
        /// Construtor using the default settings
        /// </summary>
        public WorkbookBuilder() {
            Settings = HtmlToExcelSettings.Defaults;
            Workbook = new XLWorkbook();
        }

        /// <summary>
        /// Constructor for customizing the settings
        /// </summary>
        /// <param name="settings">Settings to use for this class instance</param>
        public WorkbookBuilder(HtmlToExcelSettings settings) {
            Settings = settings;
            Workbook = new XLWorkbook();
        }

        /// <summary>
        /// Add a new sheet to the workbook that is being created
        /// </summary>
        /// <param name="sheetName">Name of the sheet</param>
        /// <param name="htmlString">HTML string to generate the the table from</param>
        /// <param name="settings">Settings for this sheet only.</param>
        /// <returns></returns>
        public WorkbookBuilder AddSheet(string sheetName, string htmlString, HtmlToExcelSettings settings = null) {
            IBrowsingContext context = BrowsingContext.New(Configuration.Default);
            IElement htmlDoc = context.OpenAsync(req => req.Content(htmlString)).Result.DocumentElement;
            IElement table = new AngleSharpUtilities().GetHtmlTableNode(htmlDoc);

            new ClosedXmlUtilities(settings ?? Settings).CreateWorksheet(Workbook, sheetName, table);
            return this;
        }

        /// <summary>
        /// Returns the current workbook as a byte array.
        /// </summary>
        /// <returns></returns>
        public byte[] GetAsByteArray() {
            using (MemoryStream stream = new MemoryStream()) {
                Workbook.SaveAs(stream);
                return stream.ToArray();
            }
        }

        /// <summary>
        /// Dispose the current Workbook
        /// </summary>
        public void Dispose() {
            Workbook.Dispose();
        }
    }
}
