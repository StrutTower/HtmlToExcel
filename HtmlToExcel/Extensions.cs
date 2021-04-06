namespace TowerSoft.HtmlToExcel {
    internal static class Extensions {
        internal static string SafeTrim(this string thisString) {
            if (!string.IsNullOrWhiteSpace(thisString)) {
                return thisString.Trim();
            }
            return string.Empty;
        }
    }
}
