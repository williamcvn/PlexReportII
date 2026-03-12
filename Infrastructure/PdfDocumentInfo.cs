// File Version: 1.0.0
// Last Modified: 2026-02-04
// Change Owner: Office of William

namespace PlexReportII.Infrastructure
{
    /// <summary>
    /// PDF 文件資訊設定類別。
    /// </summary>
    public class PdfDocumentInfo
    {
        /// <summary>
        /// 文件標題。
        /// </summary>
        public string Title { get; set; } = string.Empty;

        /// <summary>
        /// 文件作者。
        /// </summary>
        public string Author { get; set; } = "Office of William";

        /// <summary>
        /// 文件主旨。
        /// </summary>
        public string Subject { get; set; } = string.Empty;

        /// <summary>
        /// 是否允許複製內容。
        /// </summary>
        public bool AllowCopyContent { get; set; } = false;

        /// <summary>
        /// 是否允許編輯內容。
        /// </summary>
        public bool AllowEditContent { get; set; } = false;

        /// <summary>
        /// 是否允許列印。
        /// </summary>
        public bool AllowPrint { get; set; } = true;

        /// <summary>
        /// 是否允許編輯註解。
        /// </summary>
        public bool AllowEditAnnotations { get; set; } = true;
    }
}
