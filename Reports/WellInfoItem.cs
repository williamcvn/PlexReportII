// File Version: 1.0.0
// Last Modified: 2026-02-25
// Change Owner: Office of William

namespace PlexReportII.Reports
{
    /// <summary>
    /// Well Info 資料項目。
    /// 用於 DrawWellInfoTable 繪製每列 Key-Value 資訊。
    /// </summary>
    public class WellInfoItem
    {
        /// <summary>
        /// 標籤名稱（左側顯示）。
        /// </summary>
        public string Key { get; set; } = "";

        /// <summary>
        /// 標籤值。
        /// </summary>
        public string Value { get; set; } = "";

        /// <summary>
        /// 是否為雙欄佈局。
        /// true: Key 左對齊、Value 右對齊。
        /// false: 單行 "Key: Value" 左對齊。
        /// </summary>
        public bool Is2Column { get; set; } = false;
    }
}
