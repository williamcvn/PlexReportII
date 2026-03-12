// File Version: 1.0.0
// Last Modified: 2026-02-11
// Change Owner: Office of William

using System.Collections.Generic;
using System.Drawing;

namespace PlexReportII.Reports
{
    /// <summary>
    /// PC/NC Detail 資料列 (對應 API 資料或 CSV 中繼)。
    /// 每個 Instance 代表表格中的一列資料。
    /// </summary>
    public class PcncDetailItem
    {
        /// <summary>孔位標識 (A1, A2)</summary>
        public string WellId { get; set; } = string.Empty;

        /// <summary>控制類型 (Positive Control, Negative Control)</summary>
        public string Control { get; set; } = string.Empty;

        /// <summary>核苷酸變異 (可能含多行)</summary>
        public string NucleotideChange { get; set; } = string.Empty;

        /// <summary>突變描述 (可能含多行)</summary>
        public string Mutation { get; set; } = string.Empty;

        /// <summary>中位螢光強度</summary>
        public string MFI { get; set; } = string.Empty;

        /// <summary>閾值 (含逗號格式, 如 "< 4,000")</summary>
        public string Cutoff { get; set; } = string.Empty;

        /// <summary>
        /// 欄位顏色覆寫 (Key = 欄位索引, Value = 顏色)。
        /// 未包含的欄位使用預設黑色。
        /// DLL 端不含業務邏輯，僅依此設定繪製。
        /// </summary>
        public Dictionary<int, Color> CellColorOverrides { get; set; }
            = new Dictionary<int, Color>();

        /// <summary>
        /// 取得指定欄位索引的值。
        /// </summary>
        /// <param name="columnIndex">欄位索引 (0-5)</param>
        /// <returns>對應欄位的文字值</returns>
        public string GetValueByIndex(int columnIndex)
        {
            return columnIndex switch
            {
                0 => WellId,
                1 => Control,
                2 => NucleotideChange,
                3 => Mutation,
                4 => MFI,
                5 => Cutoff,
                _ => string.Empty
            };
        }
    }
}
