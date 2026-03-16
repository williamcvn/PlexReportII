// File Version: 1.0.0
// Last Modified: 2026-02-04
// Change Owner: Office of William

using System.Drawing;

namespace PlexReportII.Infrastructure
{
    /// <summary>
    /// PDF 全域設定類別。
    /// 提供集中管理 PDF 報表的頁面邊界、預設字型等設定。
    /// </summary>
    public static class PdfGlobalConfig
    {
        /// <summary>
        /// 水平邊距（左右各縮減的點數，預設 30）。
        /// </summary>
        public static float MarginHorizontal { get; set; } = 30f;

        /// <summary>
        /// 垂直邊距（上下各縮減的點數，預設 60）。
        /// </summary>
        public static float MarginVertical { get; set; } = 60f;

        public static float SpaceAreaHeight_3pt { get; set; } = 3f;

        /// <summary>
        /// 預設換頁底部預留空間 (pt)。用於控制表格換頁時，最底端距離 Footer 頂線的距離。
        /// 預設值為 30f。
        /// </summary>
        public static float DefaultPageBottomMargin { get; set; } = 30f;

        /// <summary>
        /// 取得頁面的繪圖區域（已套用邊距）。
        /// </summary>
        /// <param name="pageRectangle">原始頁面矩形</param>
        /// <returns>已縮減邊距的繪圖區域</returns>
        public static RectangleF GetContentRect(RectangleF pageRectangle)
        {
            RectangleF rcPage = pageRectangle;
            rcPage.Inflate(-MarginHorizontal, -MarginVertical);
            //Todo: 如需額外頂部空間，可取消下列註解
            //// RectangleF.Top 是唯讀，需調整 Y 與 Height
            //rcPage.Y += SpaceAreaHeight_3pt; // 頂部額外空間
            //rcPage.Height -= SpaceAreaHeight_3pt;
            return rcPage;
        }

        /// <summary>
        /// 重設為預設值。
        /// </summary>
        public static void ResetToDefaults()
        {
            MarginHorizontal = 30f;
            MarginVertical = 60f;
        }

        /// <summary>
        /// 設定自訂邊距。
        /// </summary>
        /// <param name="horizontal">水平邊距</param>
        /// <param name="vertical">垂直邊距</param>
        public static void SetMargins(float horizontal, float vertical)
        {
            MarginHorizontal = horizontal;
            MarginVertical = vertical;
        }
    }
}
