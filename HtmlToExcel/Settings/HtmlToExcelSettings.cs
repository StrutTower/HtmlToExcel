namespace TowerSoft.HtmlToExcel {
    public class HtmlToExcelSettings {
        /// <summary>
        /// Toggles if all used cells should autofit to the contents of the cell. Default = true
        /// </summary>
        public bool AutofitColumns { get; set; }

        /// <summary>
        /// Toggles if the main rows of the data should be striped. Does NOT work if the sheet contains ANY merged cells. Default = true
        /// </summary>
        public bool ShowRowStripes { get; set; }

        /// <summary>
        /// Toggles if the table should show filters in the header. Does NOT work if the sheet contains ANY merged cells. Default = true
        /// </summary>
        public bool ShowFilter { get; set; }

        /// <summary>
        /// Get the default settings
        /// </summary>
        public static HtmlToExcelSettings Defaults {
            get {
                return new HtmlToExcelSettings {
                    AutofitColumns = true,
                    ShowRowStripes = true,
                    ShowFilter = true
                };
            }
        }
    }
}
