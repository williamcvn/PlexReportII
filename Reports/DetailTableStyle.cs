// File Version: 1.0.0
// Last Modified: 2026-02-11
// Change Owner: Office of William

using System;
using System.Drawing;

namespace PlexReportII.Reports
{
    /// <summary>
    /// 模組化表格樣式定義，控制所有表格的繪製行為。
    /// 透過此類別可彈性設定欄位數、寬度、合併、奇偶列變色、格線、換頁等所有樣式參數。
    /// </summary>
    public class DetailTableStyle
    {
        // ─── 欄位定義 (動態欄位數) ───

        /// <summary>表格標題列文字 (Length 決定欄位數)</summary>
        public string[] Headers { get; set; } = Array.Empty<string>();

        /// <summary>寬度計算基礎分母 (PageWidth / BaseDivider 為一個單位)</summary>
        public int WidthBaseDivider { get; set; } = 8;

        /// <summary>每欄佔幾個基礎單位</summary>
        public float[] ColumnWidthFactors { get; set; } = Array.Empty<float>();

        /// <summary>每欄微調值 (pt), 正=加寬, 負=縮窄</summary>
        public float[] ColumnWidthOffsets { get; set; } = Array.Empty<float>();

        /// <summary>每欄對齊方式 (0=Left, 1=Right, 2=Center)</summary>
        public int[] ColumnAlignments { get; set; } = Array.Empty<int>();

        // ─── 合併控制 ───

        /// <summary>是否啟用欄位垂直合併</summary>
        public bool EnableColumnMerge { get; set; } = false;

        /// <summary>需要垂直合併的欄位索引 (僅 EnableColumnMerge=true 時有效)</summary>
        public int[] MergeColumnIndices { get; set; } = Array.Empty<int>();

        // ─── 奇偶列交替網底 ───

        /// <summary>是否啟用奇偶列交替網底</summary>
        public bool AlternatingRowBackground { get; set; } = false;

        /// <summary>偶數列背景色 (row 0, 2, 4...)</summary>
        public Color EvenRowColor { get; set; } = Color.White;

        /// <summary>奇數列背景色 (row 1, 3, 5...)</summary>
        public Color OddRowColor { get; set; } = Color.FromArgb(245, 245, 245);

        // ─── 格線樣式 ───

        /// <summary>格線粗細 (pt)</summary>
        public float BorderWidth { get; set; } = 0.2f;

        /// <summary>格線顏色</summary>
        public Color BorderColor { get; set; } = Color.Gray;

        /// <summary>列分隔線模式</summary>
        public RowSeparatorMode RowSeparator { get; set; } = RowSeparatorMode.SkipMergedColumns;

        // ─── 字型與間距 ───

        /// <summary>字型大小</summary>
        public float FontSize { get; set; } = 7f;

        /// <summary>文字內縮值 (pt)</summary>
        public float CellPadding { get; set; } = 3.5f;

        // ─── 換頁控制 ───

        /// <summary>頁面頂部預留空間 (pt)</summary>
        public float PageTopMargin { get; set; } = 20f;

        /// <summary>頁面底部預留空間 (pt)</summary>
        public float PageBottomMargin { get; set; } = 30f;

        /// <summary>是否有 Footer Note</summary>
        public bool HasFooterNote { get; set; } = false;

        /// <summary>Footer Note 預留高度 (pt)</summary>
        public float FooterNoteHeight { get; set; } = 0f;

        /// <summary>換頁後是否重繪標題列</summary>
        public bool RedrawHeaderOnNewPage { get; set; } = true;

        /// <summary>
        /// 驗證樣式參數的一致性。
        /// </summary>
        public void Validate()
        {
            if (Headers == null || Headers.Length == 0)
            {
                throw new ArgumentException("Headers 不可為空");
            }

            int colCount = Headers.Length;

            if (ColumnWidthFactors == null || ColumnWidthFactors.Length != colCount)
            {
                throw new ArgumentException($"ColumnWidthFactors 長度 ({ColumnWidthFactors?.Length}) 必須與 Headers 長度 ({colCount}) 一致");
            }

            if (ColumnWidthOffsets == null || ColumnWidthOffsets.Length != colCount)
            {
                throw new ArgumentException($"ColumnWidthOffsets 長度 ({ColumnWidthOffsets?.Length}) 必須與 Headers 長度 ({colCount}) 一致");
            }

            if (ColumnAlignments == null || ColumnAlignments.Length != colCount)
            {
                throw new ArgumentException($"ColumnAlignments 長度 ({ColumnAlignments?.Length}) 必須與 Headers 長度 ({colCount}) 一致");
            }

            if (EnableColumnMerge && (MergeColumnIndices == null || MergeColumnIndices.Length == 0))
            {
                throw new ArgumentException("啟用合併時 MergeColumnIndices 不可為空");
            }
        }
    }
}
