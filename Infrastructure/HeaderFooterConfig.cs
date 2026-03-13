// File Version: 1.0.0
// Last Modified: 2026-02-04
// Change Owner: Office of William

using System;
using System.Drawing;
using C1.Pdf;
using C1.Util;
using C1Font = C1.Util.Font;

namespace PlexReportII.Infrastructure
{
    /// <summary>
    /// Header/Footer 繪製委派。
    /// </summary>
    /// <param name="pdf">PDF 文件</param>
    /// <param name="pageIndex">頁面索引（0-based）</param>
    /// <param name="totalPages">總頁數</param>
    /// <param name="rect">繪製區域</param>
    public delegate void RenderHeaderFooterDelegate(C1PdfDocument pdf, int pageIndex, int totalPages, RectangleF rect);

    /// <summary>
    /// Header/Footer 設定類別。
    /// 提供 PDF 報表頁首/頁尾的設定與繪製邏輯。
    /// </summary>
    public class HeaderFooterConfig
    {
        /// <summary>
        /// Header 高度（點數，預設 30）。
        /// </summary>
        public float HeaderHeight { get; set; } = 30f;

        /// <summary>
        /// Header 與內容區域的間距（預設 6）。
        /// </summary>
        public float HeaderOffset { get; set; } = 6f;

        /// <summary>
        /// Footer 高度（點數，預設 12）。
        /// </summary>
        public float FooterHeight { get; set; } = 12f;

        /// <summary>
        /// Footer 與內容區域的間距（預設 6）。
        /// </summary>
        public float FooterOffset { get; set; } = 6f;

        /// <summary>
        /// 是否顯示 Header。
        /// </summary>
        public bool ShowHeader { get; set; } = true;

        /// <summary>
        /// 是否顯示 Footer。
        /// </summary>
        public bool ShowFooter { get; set; } = true;

        /// <summary>
        /// 是否顯示頁碼。
        /// </summary>
        public bool ShowPageNumber { get; set; } = true;

        /// <summary>
        /// Header 標題文字。
        /// </summary>
        public string HeaderTitle { get; set; } = string.Empty;

        /// <summary>
        /// Header Logo 圖片。
        /// </summary>
        public Image? Logo { get; set; }

        /// <summary>
        /// 軟體名稱資訊（預設 "DeXipher™"）。
        /// </summary>
        public string SoftwareNameText { get; set; } = "DeXipher™";

        /// <summary>
        /// 軟體版本資訊（預設 "1.0.0.3643"）。
        /// </summary>
        public string VersionText { get; set; } = "1.0.0.3643";

        /// <summary>
        /// 操作者名稱（預設 "William"）。
        /// </summary>
        public string OperatorName { get; set; } = "William";

        /// <summary>
        /// 匯出時間字串。
        /// </summary>
        public string ExportedTime { get; set; } = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

        /// <summary>
        /// 是否顯示 Research Use Only。
        /// </summary>
        public bool IsResearchUseOnly { get; set; } = false;

        /// <summary>
        /// Research Use Only 提示文字。
        /// </summary>
        public string ResearchUseOnlyText { get; set; } = "Research Use Only";

        /// <summary>
        /// Footer 文字（左側）。
        /// </summary>
        public string FooterText { get; set; } = string.Empty;

        /// <summary>
        /// 頁碼格式（預設 "Page {0} of {1}"）。
        /// </summary>
        public string PageNumberFormat { get; set; } = "Page {0} of {1}";

        /// <summary>
        /// 起始頁（用於略過封面頁，預設 0）。
        /// </summary>
        public int StartPage { get; set; } = 0;

        /// <summary>
        /// Header 字型。
        /// </summary>
        public C1Font HeaderFont { get; set; } = new C1Font("Arial", 17, C1.Util.FontStyle.Bold);

        /// <summary>
        /// Footer 字型。
        /// </summary>
        public C1Font FooterFont { get; set; } = new C1Font("Arial", 7, C1.Util.FontStyle.Regular);

        /// <summary>
        /// Header 文字顏色。
        /// </summary>
        public Color HeaderColor { get; set; } = Color.Black;

        /// <summary>
        /// Footer 文字顏色。
        /// </summary>
        public Color FooterColor { get; set; } = Color.Gray;

        /// <summary>
        /// 是否繪製 Header 底線。
        /// </summary>
        public bool DrawHeaderLine { get; set; } = true;

        /// <summary>
        /// 是否繪製 Footer 頂線。
        /// </summary>
        public bool DrawFooterLine { get; set; } = true;

        /// <summary>
        /// 自訂 Header 繪製邏輯（若設定則覆寫預設繪製）。
        /// </summary>
        public RenderHeaderFooterDelegate? OnRenderHeader { get; set; }

        /// <summary>
        /// 自訂 Footer 繪製邏輯（若設定則覆寫預設繪製）。
        /// </summary>
        public RenderHeaderFooterDelegate? OnRenderFooter { get; set; }

        /// <summary>
        /// 計算 Header 繪製區域。
        /// </summary>
        /// <param name="contentRect">內容區域</param>
        /// <returns>Header 區域</returns>
        public RectangleF GetHeaderRect(RectangleF contentRect)
        {
            return new RectangleF(
                contentRect.X,
                contentRect.Y - HeaderOffset - HeaderHeight,
                contentRect.Width,
                HeaderHeight);
        }

        /// <summary>
        /// 計算 Footer 繪製區域。
        /// </summary>
        /// <param name="contentRect">內容區域</param>
        /// <returns>Footer 區域</returns>
        public RectangleF GetFooterRect(RectangleF contentRect)
        {
            return new RectangleF(
                contentRect.X,
                contentRect.Bottom + FooterOffset,
                contentRect.Width,
                FooterHeight);
        }
    }
}
