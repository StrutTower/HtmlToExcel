using AngleSharp;
using AngleSharp.Dom;
using System;
using TowerSoft.HtmlToExcel.Utilities;

namespace TowerSoft.HtmlToExcel {
    /// <summary>
    /// Excel generator class for creating a Excel spreadsheet with a single sheet.
    /// </summary>
    public class WorkbookGenerator {
        /// <summary>
        /// Current HtmlToExcelSettings used by this instance
        /// </summary>
        public HtmlToExcelSettings HtmlToExcelSettings { get; }

        /// <summary>
        /// Create a new WorkbookGenerator instance with the default settings
        /// </summary>
        public WorkbookGenerator() {
            HtmlToExcelSettings = HtmlToExcelSettings.Defaults;
        }

        /// <summary>
        /// Create a new WorkbookGenerator instance with custom settings
        /// </summary>
        /// <param name="settings">Settings class with customized settings</param>
        public WorkbookGenerator(HtmlToExcelSettings settings) {
            HtmlToExcelSettings = settings;
        }

        /// <summary>
        /// Generates an Excel file from the HTML string and returns the byte array of the file data.
        /// </summary>
        /// <param name="htmlString">HTML string with only one table element. Will throw an error if there are more than one tables or if a table cannot be found.</param>
        /// <returns>Byte array of the Excel file data.</returns>
        public byte[] FromHtmlString(string htmlString) {
            IBrowsingContext context = BrowsingContext.New(Configuration.Default);
            var document = context.OpenAsync(req => req.Content(htmlString)).Result;

            return ProcessDocument(document.DocumentElement);
        }

        /// <summary>
        /// Generates an Excel file from the returned string from the supplied URI and returns the byte array of the file data.
        /// </summary>
        /// <param name="uri">URI to download the HTML string from. Will throw an error if the server cannot be reached, there are more than one tables or if a table cannot be found.</param>
        /// <returns>Byte array of the Excel file data.</returns>
        public byte[] FromUri(Uri uri) {
            IBrowsingContext context = BrowsingContext.New(Configuration.Default);
            var document = context.OpenAsync(uri.ToString()).Result;
            return ProcessDocument(document.DocumentElement);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="htmlDoc"></param>
        /// <returns></returns>
        private byte[] ProcessDocument(IElement htmlDoc) {
            IElement table = new AngleSharpUtilities().GetHtmlTableNode(htmlDoc);
            return new ClosedXmlUtilities(HtmlToExcelSettings).GenerateWorkbookFromHtmlNode(table);
        }
    }
}
