// File Version: 1.0.0
// Last Modified: 2026-02-05
// Change Owner: Office of William

using System.Drawing;
using C1Font = C1.Util.Font;

namespace PlexReportII.Infrastructure
{
    /// <summary>
    /// 表格樣式設定類別 (V1 簡化版)。
    /// 用於定義 RenderTable_V1 方法的表格外觀。
    /// </summary>
    public class TbSetting
    {
        /// <summary>
        /// 標題列樣式。
        /// </summary>
        public enum HeaderStyle
        {
            /// <summary>無標題列</summary>
            NoHeader,
            /// <summary>帶底線的標題列</summary>
            HasUnderline
        }

        /// <summary>
        /// 資料列樣式。
        /// </summary>
        public enum RowStyle
        {
            /// <summary>無框線</summary>
            NoRowBorder,
            /// <summary>有框線</summary>
            HasRowBorder
        }

        /// <summary>
        /// 標題列樣式設定。
        /// </summary>
        public HeaderStyle HeaderType { get; set; } = HeaderStyle.NoHeader;

        /// <summary>
        /// 資料列樣式設定。
        /// </summary>
        public RowStyle RowType { get; set; } = RowStyle.NoRowBorder;

        /// <summary>
        /// 各欄位寬度比例 (欄位名稱 → 比例值)。
        /// </summary>
        public Dictionary<string, float> ColumnWidthFactors { get; set; } = new Dictionary<string, float>();

        /// <summary>
        /// 欄位合併設定 (起始欄位索引 → 合併欄位數)。
        /// </summary>
        public Dictionary<int, int> MergeColumns { get; set; } = new Dictionary<int, int>();

        /// <summary>
        /// 資料列字型 (使用 C1.Util.Font)。
        /// </summary>
        public C1Font? RowFont { get; set; }

        /// <summary>
        /// 標題列字型 (使用 C1.Util.Font)。
        /// </summary>
        public C1Font? HeaderFont { get; set; }

        /// <summary>
        /// 標題列文字格式。
        /// </summary>
        public StringFormat? HeaderStringFormat { get; set; }

        /// <summary>
        /// 資料列文字格式。
        /// </summary>
        public StringFormat? RowStringFormat { get; set; }

        /// <summary>
        /// 標題列高度額外邊距。
        /// </summary>
        public float HeaderHeightMargin { get; set; } = 0f;

        /// <summary>
        /// 資料列高度額外邊距。
        /// </summary>
        public float RowHeightMargin { get; set; } = 0f;

        /// <summary>
        /// 建立預設的 TbSetting1 樣式 (4 欄等寬, 地址列合併)。
        /// </summary>
        /// <param name="rowFont">資料列字型 (C1.Util.Font)</param>
        /// <returns>TbSetting1 樣式設定</returns>
        public static TbSetting CreateTbSetting1(C1Font rowFont)
        {
            StringFormat sf = new StringFormat
            {
                LineAlignment = StringAlignment.Center,
                Alignment = StringAlignment.Near
            };

            return new TbSetting
            {
                HeaderType = HeaderStyle.NoHeader,
                RowType = RowStyle.NoRowBorder,
                RowFont = rowFont,
                HeaderStringFormat = new StringFormat(sf),
                RowStringFormat = new StringFormat(sf),
                ColumnWidthFactors = new Dictionary<string, float>
                {
                    { "1", 1f },
                    { "2", 1f },
                    { "3", 1f },
                    { "4", 1f }
                },
                MergeColumns = new Dictionary<int, int>
                {
                    { 1, 3 } // 從欄位 1 開始，合併 3 個欄位 (用於地址列)
                }
            };
        }
    }
}


