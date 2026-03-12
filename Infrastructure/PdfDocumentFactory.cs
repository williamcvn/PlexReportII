// File Version: 1.0.0
// Last Modified: 2026-02-04
// Change Owner: Office of William

using System;
using C1.Pdf;

namespace PlexReportII.Infrastructure
{
    /// <summary>
    /// PDF 文件工廠類別，負責建立與初始化 C1PdfDocument 實例。
    /// Implements SRS-PDF-001: PDF Document Creation
    /// </summary>
    public static class PdfDocumentFactory
    {
        /// <summary>
        /// 建立一個標準 A4 格式的 PDF 文件。
        /// </summary>
        /// <param name="documentInfo">文件資訊設定</param>
        /// <returns>已初始化的 C1PdfDocument 實例</returns>
        public static C1PdfDocument CreateDocument(PdfDocumentInfo documentInfo)
        {
            if (documentInfo == null)
            {
                throw new ArgumentNullException(nameof(documentInfo));
            }

            C1PdfDocument pdf = new C1PdfDocument();

            // 基本設定
            pdf.Landscape = false;
            pdf.PaperKind = GrapeCity.Documents.Common.PaperKind.A4;

            // 文件資訊
            pdf.DocumentInfo.Title = documentInfo.Title ?? string.Empty;
            pdf.DocumentInfo.Author = documentInfo.Author ?? "Office of William";
            pdf.DocumentInfo.Subject = documentInfo.Subject ?? string.Empty;

            // 安全性設定
            pdf.Security.AllowCopyContent = documentInfo.AllowCopyContent;
            pdf.Security.AllowEditContent = documentInfo.AllowEditContent;
            pdf.Security.AllowPrint = documentInfo.AllowPrint;
            pdf.Security.AllowEditAnnotations = documentInfo.AllowEditAnnotations;

            return pdf;
        }

        /// <summary>
        /// 使用預設設定建立 PDF 文件。
        /// </summary>
        /// <param name="title">文件標題</param>
        /// <returns>已初始化的 C1PdfDocument 實例</returns>
        public static C1PdfDocument CreateDocument(string title)
        {
            return CreateDocument(new PdfDocumentInfo
            {
                Title = title,
                Author = "Office of William",
                Subject = "Kit Report",
                AllowCopyContent = false,
                AllowEditContent = false,
                AllowPrint = true,
                AllowEditAnnotations = true
            });
        }
    }
}
