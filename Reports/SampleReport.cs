// File Version: 1.0.0
// Last Modified: 2026-02-04
// Change Owner: Office of William

using System.Drawing;
using System.Collections.Generic;
using C1.Util;
using PlexReportII.Infrastructure;
using C1Font = C1.Util.Font;

namespace PlexReportII.Reports
{
    /// <summary>
    /// 範例 PDF 報表，用於測試基本功能。
    /// </summary>
    public class SampleReport : BasePdfReport
    {
        /// <summary>
        /// 報表標題。
        /// </summary>
        public override string ReportTitle => "Sample Report";

        /// <summary>
        /// 初始化範例報表。
        /// </summary>
        /// <param name="logger">日誌服務</param>
        public SampleReport(IPlexLogger? logger = null) : base(logger)
        {
        }

        /// <summary>
        /// 繪製報表內容。
        /// </summary>
        protected override void RenderContent()
        {
            // 繪製標題
            DrawString("Sample Report - PlexReportII", TitleFont, Color.Black);

            // 空行
            CurrentRect = new RectangleF(
                CurrentRect.X,
                CurrentRect.Y + 20,
                CurrentRect.Width,
                CurrentRect.Height - 20);

            // 繪製內文
            DrawString("此為 PlexReportII 專案的範例報表，使用 ComponentOne C1.Pdf 產生。", BodyFont, Color.DarkGray);

            // 繪製說明
            CurrentRect = new RectangleF(
                CurrentRect.X,
                CurrentRect.Y + 10,
                CurrentRect.Width,
                CurrentRect.Height - 10);

            DrawString("作者: Office of William", NoteFont, Color.Gray);
            DrawString($"產生時間: {System.DateTime.Now:yyyy-MM-dd HH:mm:ss}", NoteFont, Color.Gray);
            DrawString($"產生時間: {System.DateTime.Now:yyyy-MM-dd HH:mm:ss}", NoteFont, Color.Gray);
        }

        /// <summary>
        /// 繪製多色段落。
        /// 輸入格式範例: "紅色文字|Red;藍色文字|Blue;一般文字" (若未指定顏色預設為黑色)
        /// </summary>
        /// <param name="inputFormat">格式化字串</param>
        /// <param name="outline">是否加入書籤 (Outline)</param>
        /// <param name="linkTarget">是否設定為連結目標 (LinkTarget)</param>
        public void DrawMultiColorParagraph(string inputFormat, bool outline = false, bool linkTarget = false)
        {
            if (string.IsNullOrWhiteSpace(inputFormat))
            {
                return;
            }

            var strTokens = new List<string>();
            var colorTokens = new List<Color>();

            // 解析字串 (以分號分隔區段)
            string[] segments = inputFormat.Split(';');
            
            foreach (string segment in segments)
            {
                if (string.IsNullOrWhiteSpace(segment)) continue;

                string text = segment;
                string colorName = "Black";

                // 解析顏色 (以直線分隔文字與顏色)
                int pipeIndex = segment.IndexOf('|');
                if (pipeIndex >= 0)
                {
                    text = segment.Substring(0, pipeIndex);
                    if (pipeIndex < segment.Length - 1)
                    {
                        colorName = segment.Substring(pipeIndex + 1).Trim();
                    }
                }

                strTokens.Add(text);
                
                // 轉換顏色
                Color color = Color.FromName(colorName);
                if (color.A == 0 && color.R == 0 && color.G == 0 && color.B == 0 && colorName != "Black")
                {
                    // 若無法辨識顏色名稱，預設為黑色
                    color = Color.Black; 
                }
                colorTokens.Add(color);
            }

            if (strTokens.Count > 0)
            {
                // 組合完整文字用於計算高度 (僅作參考，RenderParagraph_Color 內部會重新計算)
                string fullText = string.Join("", strTokens);
                
                // 呼叫 BasePdfReport 的 RenderParagraph_Color
                CurrentRect = RenderParagraph_Color(
                    Pdf, 
                    fullText, 
                    strTokens, 
                    colorTokens, 
                    BodyFont, // 使用內文字型
                    PageRect, 
                    CurrentRect, 
                    outline, // Outline
                    linkTarget // LinkTarget
                );
            }
        }

        /// <summary>
        /// 繪製 PC/NC 註解 (從 CSV 載入的項目)。
        /// </summary>
        /// <param name="items">註解項目列表</param>
        public void DrawPcncNote(List<PcncLegendItem> items)
        {
            if (items == null || items.Count == 0) return;

            var lines = new List<string>();
            foreach (var item in items)
            {
                if (string.IsNullOrWhiteSpace(item.ItemText)) continue;

                // 處理 ItemText 中的換行符號
                var itemLines = item.ItemText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                lines.AddRange(itemLines);
            }

            if (lines.Count > 0)
            {
                // 使用 NoteFont 的大小，預設為 8
                float fontSize = NoteFont != null ? NoteFont.Size : 8;
                CurrentRect = RenderPcncNote(Pdf, lines, PageRect, CurrentRect, fontSize);
            }
        }

        /// <summary>
        /// 繪製 PC/NC Table (從 CSV 載入的項目)。
        /// </summary>
        /// <param name="items">Table 項目列表</param>
        public void DrawPcncTable(List<PcncTableItem> items)
        {
            if (items == null || items.Count == 0) return;

            // 使用 NoteFont (or generic font) size
            float fontSize = NoteFont != null ? NoteFont.Size : 8;
            CurrentRect = RenderPcncTable(Pdf, items, PageRect, CurrentRect, fontSize);
        }

        /// <summary>
        /// 繪製 PC/NC Fail Detail Table (從 CSV 或 API 載入的項目)。
        /// </summary>
        /// <param name="items">資料列表 (扁平, 已排序)</param>
        /// <param name="style">表格樣式定義</param>
        public void DrawPcncDetailTable(List<PcncDetailItem> items, DetailTableStyle style)
        {
            if (items == null || items.Count == 0) return;

            CurrentRect = RenderPcncDetailTable(Pdf, items, style, PageRect, CurrentRect);
        }
    }
}
