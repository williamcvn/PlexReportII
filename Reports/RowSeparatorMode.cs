// File Version: 1.0.0
// Last Modified: 2026-02-11
// Change Owner: Office of William

namespace PlexReportII.Reports
{
    /// <summary>
    /// 列分隔線模式。
    /// </summary>
    public enum RowSeparatorMode
    {
        /// <summary>完整寬度 (Left → Right)</summary>
        FullWidth,

        /// <summary>自動跳過合併欄 (第一個非合併欄 → Right)</summary>
        SkipMergedColumns,

        /// <summary>不繪製列分隔線</summary>
        None
    }
}
