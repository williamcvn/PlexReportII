// File Version: 1.0.0
// Last Modified: 2026-02-04
// Change Owner: Office of William

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;

namespace PlexReportII.Infrastructure
{
    /// <summary>
    /// 日誌服務實作類別。
    /// 支援 Console 彩色輸出、檔案寫入、日期分割與錯誤位置追蹤。
    /// </summary>
    public class PlexLogger : IPlexLogger
    {
        private readonly string _logDirectory;
        private readonly object _lockObject = new object();
        private readonly long _maxFileSizeBytes;
        private string _currentLogFilePath = string.Empty;
        private DateTime _currentLogDate;

        /// <summary>
        /// 初始化日誌服務。
        /// </summary>
        /// <param name="logDirectory">日誌存放目錄（預設為執行路徑下的 Log 資料夾）</param>
        /// <param name="maxFileSizeMB">單一日誌檔案最大大小（MB），超過則分割</param>
        public PlexLogger(string? logDirectory = null, int maxFileSizeMB = 10)
        {
            // 依據 spec.md 規定：發行 DLL 在運行時，log 存入執行路徑下的 Log 資料夾
            _logDirectory = logDirectory ?? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Log");
            _maxFileSizeBytes = maxFileSizeMB * 1024L * 1024L;

            if (!Directory.Exists(_logDirectory))
            {
                Directory.CreateDirectory(_logDirectory);
            }

            _currentLogDate = DateTime.Today;
            UpdateLogFilePath();
        }

        /// <summary>
        /// 記錄除錯訊息。
        /// </summary>
        public void Debug(string message)
        {
            Log(LogType.Debug, message);
        }

        /// <summary>
        /// 記錄一般資訊。
        /// </summary>
        public void Info(string message)
        {
            Log(LogType.Info, message);
        }

        /// <summary>
        /// 記錄警告訊息。
        /// </summary>
        public void Warn(string message)
        {
            Log(LogType.Warn, message);
        }

        /// <summary>
        /// 記錄錯誤訊息（含錯誤位置資訊）。
        /// </summary>
        public void Error(string message, Exception? exception = null,
            [CallerFilePath] string filePath = "",
            [CallerLineNumber] int lineNumber = 0)
        {
            string locationInfo = $"[{Path.GetFileName(filePath)}:{lineNumber}]";
            string fullMessage = $"{locationInfo} {message}";

            if (exception != null)
            {
                fullMessage += $"\n例外: {exception.GetType().Name}: {exception.Message}\n{exception.StackTrace}";
            }

            Log(LogType.Error, fullMessage);
        }

        /// <summary>
        /// 記錄訊息（僅寫入檔案，不顯示於 Console）。
        /// </summary>
        public void Trace(string message)
        {
            WriteToFile(LogType.Debug, message);
        }

        private void Log(LogType logType, string message)
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string formattedMessage = $"[{timestamp}] [{logType}] {message}";

            // 寫入 Console（依類型顯示不同顏色）
            WriteToConsole(logType, formattedMessage);

            // 寫入檔案
            WriteToFile(logType, formattedMessage);
        }

        private void WriteToConsole(LogType logType, string message)
        {
            ConsoleColor originalColor = Console.ForegroundColor;

            Console.ForegroundColor = logType switch
            {
                LogType.Debug => ConsoleColor.Gray,
                LogType.Info => ConsoleColor.White,
                LogType.Warn => ConsoleColor.Yellow,
                LogType.Error => ConsoleColor.Red,
                _ => ConsoleColor.White
            };

            Console.WriteLine(message);
            Console.ForegroundColor = originalColor;
        }

        private void WriteToFile(LogType logType, string message)
        {
            lock (_lockObject)
            {
                // 檢查日期是否變更（需要新的日誌檔）
                if (DateTime.Today != _currentLogDate)
                {
                    _currentLogDate = DateTime.Today;
                    UpdateLogFilePath();
                }

                // 檢查檔案大小是否超過限制
                if (File.Exists(_currentLogFilePath))
                {
                    FileInfo fileInfo = new FileInfo(_currentLogFilePath);
                    if (fileInfo.Length >= _maxFileSizeBytes)
                    {
                        UpdateLogFilePath(true);
                    }
                }

                try
                {
                    File.AppendAllText(_currentLogFilePath, message + Environment.NewLine, Encoding.UTF8);
                }
                catch (Exception ex)
                {
                    // 日誌寫入失敗時，僅輸出到 Console 避免無限迴圈
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine($"[LOG WRITE ERROR] {ex.Message}");
                    Console.ResetColor();
                }
            }
        }

        private void UpdateLogFilePath(bool forceNewFile = false)
        {
            string dateStr = _currentLogDate.ToString("yyyyMMdd");
            string baseFileName = $"PlexReportII_{dateStr}";

            if (forceNewFile)
            {
                // 尋找可用的序號
                int sequence = 1;
                string newPath;
                do
                {
                    newPath = Path.Combine(_logDirectory, $"{baseFileName}_{sequence:D3}.log");
                    sequence++;
                } while (File.Exists(newPath) && new FileInfo(newPath).Length >= _maxFileSizeBytes);

                _currentLogFilePath = newPath;
            }
            else
            {
                _currentLogFilePath = Path.Combine(_logDirectory, $"{baseFileName}.log");
            }
        }
    }
}
