using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;

namespace TowerSoft.HtmlToExcel {
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
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(htmlString);
            return ProcessHtmlDocument(htmlDoc);
        }

        /// <summary>
        /// Generates an Excel file from the returned string from the supplied URI and returns the byte array of the file data.
        /// </summary>
        /// <param name="uri">URI to download the HTML string from. Will throw an error if the server cannot be reached, there are more than one tables or if a table cannot be found.</param>
        /// <returns>Byte array of the Excel file data.</returns>
        public byte[] FromUri(Uri uri) {
            string html;
            using (WebClient webClient = new WebClient() { UseDefaultCredentials = true }) {
                html = webClient.DownloadString(uri);
            }
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(html);
            return ProcessHtmlDocument(htmlDoc);
        }

        /// <summary>
        /// Generates an Excel file from HTML document stream and returns the byte array of the file data.
        /// </summary>
        /// <param name="stream">Stream of the HTML document. Will throw an error if there are more than one tables or if a table cannot be found.</param>
        /// <returns>Byte array of the Excel file data.</returns>
        public byte[] FromStream(System.IO.Stream stream) {
            HtmlDocument htmlDoc = new HtmlDocument();
            htmlDoc.Load(stream);
            return ProcessHtmlDocument(htmlDoc);
        }

        private byte[] ProcessHtmlDocument(HtmlDocument htmlDoc) {
            HtmlNode table = new HtmlAgilityUtilities().GetHtmlTableNode(htmlDoc);
            return new EPPlusUtilities(HtmlToExcelSettings).GenerateWorkbookFromHtmlNode(table);
        }
    }
}
