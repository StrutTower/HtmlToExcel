using AngleSharp;
using OfficeOpenXml;
using System;

namespace TowerSoft.HtmlToExcel {
    /// <summary>
    /// Used to build a workbook with multiple sheets. Make sure to dispose of this class.
    /// </summary>
    public class WorkbookBuilder : IDisposable {
        private HtmlToExcelSettings Settings { get; }
        private ExcelPackage Package { get; }

        /// <summary>
        /// Construtor using the default settings
        /// </summary>
        public WorkbookBuilder() {
            Settings = HtmlToExcelSettings.Defaults;
            Package = new ExcelPackage();
        }

        /// <summary>
        /// Constructor for customizing the settings
        /// </summary>
        /// <param name="settings">Settings to use for this class instance</param>
        public WorkbookBuilder(HtmlToExcelSettings settings) {
            Settings = settings;
            Package = new ExcelPackage();
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
            var document = context.OpenAsync(req => req.Content(htmlString)).Result;

            new EPPlusUtilities(settings ?? Settings).CreateSheet(Package, sheetName, document.DocumentElement);
            return this;
        }

        /// <summary>
        /// Returns the current workbook as a byte array.
        /// </summary>
        /// <returns></returns>
        public byte[] GetAsByteArray() {
            return Package.GetAsByteArray();
        }

        /// <summary>
        /// 
        /// </summary>
        public void Dispose() {
            Package.Dispose();
        }
    }
}
