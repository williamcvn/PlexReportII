// File Version: 1.0.0
// Last Modified: 2026-02-04
// Change Owner: Office of William

using System.Threading.Tasks;

namespace PlexReportII.Abstractions
{
    /// <summary>
    /// 報表產生器介面。
    /// </summary>
    public interface IReportGenerator
    {
        /// <summary>
        /// 非同步產生報表並儲存至指定路徑。
        /// </summary>
        /// <param name="outputPath">輸出路徑</param>
        /// <returns>非同步任務</returns>
        Task GenerateAsync(string outputPath);

        /// <summary>
        /// 取得報表標題。
        /// </summary>
        string ReportTitle { get; }
    }
}
