// File Version: 1.0.0
// Last Modified: 2026-02-04
// Change Owner: Office of William

using System;
using System.Runtime.CompilerServices;

namespace PlexReportII.Infrastructure
{
    /// <summary>
    /// 日誌訊息類型列舉。
    /// </summary>
    public enum LogType
    {
        /// <summary>除錯訊息</summary>
        Debug,
        /// <summary>一般資訊</summary>
        Info,
        /// <summary>警告訊息</summary>
        Warn,
        /// <summary>錯誤訊息</summary>
        Error
    }

    /// <summary>
    /// 日誌服務介面。
    /// </summary>
    public interface IPlexLogger
    {
        /// <summary>
        /// 記錄除錯訊息。
        /// </summary>
        void Debug(string message);

        /// <summary>
        /// 記錄一般資訊。
        /// </summary>
        void Info(string message);

        /// <summary>
        /// 記錄警告訊息。
        /// </summary>
        void Warn(string message);

        /// <summary>
        /// 記錄錯誤訊息（含錯誤位置資訊）。
        /// </summary>
        void Error(string message, Exception? exception = null,
            [CallerFilePath] string filePath = "",
            [CallerLineNumber] int lineNumber = 0);

        /// <summary>
        /// 記錄訊息（僅寫入檔案，不顯示於 Console）。
        /// </summary>
        void Trace(string message);
    }
}
