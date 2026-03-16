// File Version: 1.1.0
// Last Modified: 2026-02-04
// Change Owner: Office of William

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using C1.Pdf;
using C1.Util;
using PlexReportII.Abstractions;
using PlexReportII.Infrastructure;
using C1Font = C1.Util.Font;
using C1StringFormat = C1.Util.StringFormat;
using GcPen = GrapeCity.Documents.Drawing.Pen;
using GcImage = GrapeCity.Documents.Drawing.Image;

namespace PlexReportII.Reports
{
    /// <summary>
    /// PDF 報表基底類別，提供共用的繪圖方法與頁面管理。
    /// 支援 Header/Footer 延遲繪製機制。
    /// </summary>
    public abstract class BasePdfReport : IReportGenerator, IDisposable
    {
        /// <summary>
        /// C1PdfDocument 實例。
        /// </summary>
        protected C1PdfDocument Pdf { get; private set; }

        /// <summary>
        /// 頁面繪圖區域。
        /// </summary>
        public RectangleF PageRect { get; private set; }

        /// <summary>
        /// 目前繪圖位置。
        /// </summary>
        protected RectangleF CurrentRect { get; set; }

        /// <summary>
        /// 日誌服務。
        /// </summary>
        protected IPlexLogger Logger { get; }

        /// <summary>
        /// 報表標題。
        /// </summary>
        public abstract string ReportTitle { get; }

        /// <summary>
        /// Header/Footer 設定。
        /// </summary>
        protected HeaderFooterConfig HeaderFooter { get; set; } = new HeaderFooterConfig();

        /// <summary>
        /// 當前繪製 Y 座標。
        /// </summary>
        public float CurrentY => CurrentRect.Y;

        /// <summary>
        /// 設定 PDF 是否允許複製內容 (Security.AllowCopyContent)。
        /// </summary>
        public bool AllowCopyContent
        {
            get => Pdf?.Security.AllowCopyContent ?? false;
            set
            {
                if (Pdf != null)
                {
                    Pdf.Security.AllowCopyContent = value;
                }
            }
        }

        /// <summary>
        /// 當前繪製 X 座標。
        /// </summary>
        public float CurrentX => CurrentRect.X;

        /// <summary>
        /// Flag Note 預設頂部間距高度 (預設: 2pt)。
        /// </summary>
        public float FlagNoteSpacing { get; set; } = 2f;

        /// <summary>
        /// Header 佔用區域。
        /// </summary>
        public RectangleF HeaderAreaRect { get; protected set; }

        /// <summary>
        /// Footer 佔用區域。
        /// </summary>
        public RectangleF FooterAreaRect { get; protected set; }

        /// <summary>
        /// 內容繪製區域（扣除 Header/Footer 後）。
        /// </summary>
        public RectangleF ContentAreaRect { get; protected set; }

        /// <summary>
        /// 總頁數。
        /// </summary>
        public int PageCount => Pdf?.Pages.Count ?? 0;

        /// <summary>
        /// 當前頁面索引 (0-based)。
        /// </summary>
        public int CurrentPageIndex => Pdf?.CurrentPage ?? 0;

        /// <summary>
        /// 頁面邊距（已棄用，請使用 PdfGlobalConfig 設定）。
        /// </summary>
        [System.Obsolete("請使用 PdfGlobalConfig.MarginHorizontal/MarginVertical 設定邊距")]
        protected float PageMargin { get; set; } = 72f;

        // 常用字型定義（使用 C1.Util.Font）
        /// <summary>標題字型</summary>
        protected C1Font TitleFont { get; set; } = new C1Font("Arial", 14, C1.Util.FontStyle.Bold);

        /// <summary>
        /// 取得 PDF 頁面大小設定 (PaperKind 名稱)。
        /// </summary>
        public string PagePaperKind => Pdf != null ? Pdf.PaperKind.ToString() : "Custom";

        /// <summary>標頭字型</summary>
        protected C1Font TableHeaderFont { get; set; } = new C1Font("Arial", 10, C1.Util.FontStyle.Bold);

        // 快取的 GcImage 與 Stream，確保生命週期覆蓋整個 PDF 產生過程
        private GcImage? _gcLogo;
        private MemoryStream? _logoMs;
        private bool _isExported = false;

        /// <summary>內文字型</summary>
        protected C1Font BodyFont { get; set; } = new C1Font("Arial", 9, C1.Util.FontStyle.Regular);

        /// <summary>註解字型</summary>
        protected C1Font NoteFont { get; set; } = new C1Font("Arial", 8, C1.Util.FontStyle.Regular);

        /// <summary>
        /// 初始化報表基底類別。
        /// </summary>
        /// <param name="logger">日誌服務（可選）</param>
        protected BasePdfReport(IPlexLogger? logger = null)
        {
            Logger = logger ?? new PlexLogger();
            Pdf = new C1PdfDocument();

            // 設定預設 Header 標題為報表標題
            HeaderFooter.HeaderTitle = ReportTitle;
        }

        /// <summary>
        /// 設定 Header/Footer 組態。
        /// </summary>
        /// <param name="configure">設定動作</param>
        public void SetHeaderFooter(Action<HeaderFooterConfig> configure)
        {
            configure(HeaderFooter);
        }

        /// <summary>
        /// 初始化 PDF 文件。
        /// </summary>
        protected virtual void InitializeDocument()
        {
            Pdf = PdfDocumentFactory.CreateDocument(ReportTitle);
            _isExported = false;
            
            // 明確建立第一頁，確保座標系統初始化 (避免 PageRectangle 回傳哨兵值)
            if (Pdf.Pages.Count == 0)
            {
                Pdf.NewPage();
            }

            // 使用 PdfGlobalConfig 取得已套用邊距的內容區域
            PageRect = PdfGlobalConfig.GetContentRect(Pdf.PageRectangle);
            CurrentRect = PageRect;

            // 更新 Header 標題
            if (string.IsNullOrEmpty(HeaderFooter.HeaderTitle))
            {
                HeaderFooter.HeaderTitle = ReportTitle;
            }

            Logger.Info($"已初始化 PDF 文件: {ReportTitle}");
        }

        /// <summary>
        /// 用於計算高度時的額外緩衝值 (參考自 RenderParagraph_Color)
        /// </summary>
        protected const float pdfAdditionHeight = 0;

        /// <summary>
        /// 以多種顏色繪製文字段落 (Base)
        /// </summary>
        protected RectangleF RenderParagraph_Color(C1PdfDocument _c1pdf, string text, List<string> strTokens, List<Color> colorTokens, C1Font font, RectangleF rcPage, RectangleF rc, bool outline)
        {
            return RenderParagraph_Color(_c1pdf, text, strTokens, colorTokens, font, rcPage, rc, outline, false);
        }

        /// <summary>
        /// 以多種顏色繪製文字段落 (Main)
        /// </summary>
        protected RectangleF RenderParagraph_Color(C1PdfDocument _c1pdf, string text, List<string> strTokens, List<Color> colorTokens, C1Font font, RectangleF rcPage, RectangleF rc, bool outline, bool linkTarget)
        {
            // 計算總高度 (參考原始碼邏輯)
            rc.Height = _c1pdf.MeasureString(text, font, rc.Width).Height + pdfAdditionHeight;
            
            // 檢查是否需要換頁
            if (rc.Bottom > rcPage.Bottom)
            {
                _c1pdf.NewPage();
                rc.Y = rcPage.Top;
            }

            if (strTokens.Count == colorTokens.Count)
            {
                RectangleF subRC = rc;
                // 重設寬度為 0，從左側開始 (實際上 subRC 一開始是 rc，width 是整行寬度，這裡需要動態計算)
                // 參考程式碼邏輯是利用 subRC.Offset 來移動 X
                // 但 subRC.Width 被用來測量 Token 寬度，所以要先由 Token 決定
                
                for (int i = 0; i < strTokens.Count; i++)
                {
                    string token = strTokens[i];
                    // 參考程式碼：replace space with 'a' for measurement?
                    // 這裡照抄參考邏輯，雖然可能影響空白寬度計算，但保持行為一致
                    string measureToken = token.Replace(' ', 'a'); 
                    subRC.Width = _c1pdf.MeasureString(measureToken, font).Width;

                    // 檢查換行 (Wrap check)
                    // 參考程式碼: if(subRC.X + subRC.Width > rc.Width)
                    // 修正邏輯: 應檢查是否超過右邊界 (rc.Right)
                    // 但若 rc.Width 代表的是 column width 且 X 從 0 開始?? 
                    // 在此專案 CurrentRect X 是有值的 (Margin)。
                    // 所以應使用 subRC.Right > rc.Right
                    if (subRC.X + subRC.Width > rc.Right) 
                    {
                        subRC.Y += subRC.Height; // 往下移一行 (高度應為行高? 這裡直接用 subRC.Height 即當前高度，若只有單行則沒問題)
                        // 若是多行，subRC.Height 可能是總高度? 
                        // _c1pdf.MeasureString(token) 回傳的是該 token 的高度。
                        // 所以這裡是合理的。
                        
                        subRC.X = rc.X; // 回到左側
                    }

                    _c1pdf.DrawString(strTokens[i], font, colorTokens[i], subRC);

                    // 移動 X 座標至下一個 Token 起點
                    subRC.Offset(subRC.Width, 0);
                }
            }

            // Outline (Bookmark)
            if (outline)
            {
                // 使用 Pdf.AddBookmark 新增書籤，層級預設為 0
                // 注意: Bookmark 在 C1Pdf 中通常指向某個特定位置
                _c1pdf.AddBookmark(text, 0, rc.Y);
            }

            // LinkTarget (Hyperlink Target)
            if (linkTarget)
            {
                _c1pdf.AddTarget(text, rc);
            }
            
            rc.Offset(0, rc.Height);
            return rc;
        }

        /// <summary>
        /// 新增頁面。
        /// </summary>
        protected virtual void AddNewPage()
        {
            Pdf.NewPage();
            // 使用 PdfGlobalConfig 取得已套用邊距的內容區域
            PageRect = PdfGlobalConfig.GetContentRect(Pdf.PageRectangle);
            CurrentRect = PageRect;

            Logger.Debug($"新增頁面: 第 {Pdf.Pages.Count} 頁");
        }

        /// <summary>
        /// 繪製文字並更新目前位置。
        /// </summary>
        /// <param name="text">文字內容</param>
        /// <param name="font">字型</param>
        /// <param name="color">顏色</param>
        /// <param name="stringFormat">字串格式</param>
        /// <returns>繪製的高度</returns>
        protected virtual float DrawString(string text, C1Font font, Color color, C1StringFormat? stringFormat = null)
        {
            stringFormat ??= new C1StringFormat();
            
            // DrawString returns number of characters drawn, NOT height.
            // We must MeasureString to get the actual height.
            SizeF size = Pdf.MeasureString(text, font, CurrentRect.Width, stringFormat);
            float height = size.Height;
            
            Pdf.DrawString(text, font, color, CurrentRect, stringFormat);
            
            CurrentRect = new RectangleF(
                CurrentRect.X,
                CurrentRect.Y + height,
                CurrentRect.Width,
                CurrentRect.Height - height);
                
            return height;
        }

        /// <summary>
        /// 檢查剩餘空間是否足夠，不足則換頁。
        /// </summary>
        /// <param name="requiredHeight">需要的高度</param>
        /// <returns>是否已換頁</returns>
        protected virtual bool EnsureSpace(float requiredHeight)
        {
            if (CurrentRect.Height < requiredHeight)
            {
                AddNewPage();
                return true;
            }
            return false;
        }

        /// <summary>
        /// 繪製所有頁面的 Header。
        /// 在所有內容繪製完成後呼叫。
        /// </summary>
        protected virtual void RenderHeaders()
        {
            if (!HeaderFooter.ShowHeader)
            {
                return;
            }

            int totalPages = Pdf.Pages.Count;
            int savedPage = Pdf.CurrentPage;

            // 準備 Logo 圖片 (若尚未準備)
            if (HeaderFooter.Logo != null && _gcLogo == null)
            {
                _logoMs = new MemoryStream();
                HeaderFooter.Logo.Save(_logoMs, System.Drawing.Imaging.ImageFormat.Png);
                _logoMs.Position = 0;
                _gcLogo = GcImage.FromStream(_logoMs);
            }

            for (int page = HeaderFooter.StartPage; page < totalPages; page++)
            {
                Pdf.CurrentPage = page;
                RectangleF contentRect = PdfGlobalConfig.GetContentRect(Pdf.PageRectangle);
                RectangleF headerRect = HeaderFooter.GetHeaderRect(contentRect);

                // 若有自訂繪製邏輯則使用
                if (HeaderFooter.OnRenderHeader != null)
                {
                    HeaderFooter.OnRenderHeader(Pdf, page, totalPages, headerRect);
                }
                else
                {
                    // 預設繪製邏輯
                    RenderDefaultHeader(page, totalPages, contentRect, headerRect);
                }
            }

            // 還原到原本頁面
            Pdf.CurrentPage = savedPage;
            Logger.Debug($"已繪製 Header 至 {totalPages - HeaderFooter.StartPage} 頁");
        }

        /// <summary>
        /// 繪製預設 Header。
        /// </summary>
        protected virtual void RenderDefaultHeader(int pageIndex, int totalPages, RectangleF contentRect, RectangleF headerRect)
        {
            // 繪製 Logo（左側）
            if (_gcLogo != null)
            {
                RectangleF logoRect = headerRect;
                logoRect.Width = 100; // 限制 Logo 寬度
                // 保持比例繪製，靠左對齊
                Pdf.DrawImage(_gcLogo, logoRect, C1.Util.ContentAlignment.MiddleLeft, ImageSizeMode.Scale);
            }

            // 繪製標題（靠右靠下）
            if (!string.IsNullOrEmpty(HeaderFooter.HeaderTitle))
            {
                C1StringFormat sfRight = new C1StringFormat
                {
                    Alignment = C1.Util.HorizontalAlignment.Right,
                    LineAlignment = C1.Util.VerticalAlignment.Bottom
                };
                Pdf.DrawString(HeaderFooter.HeaderTitle, HeaderFooter.HeaderFont, HeaderFooter.HeaderColor, headerRect, sfRight);
            }

            // 繪製底線
            if (HeaderFooter.DrawHeaderLine)
            {
                GcPen pen = new GcPen(Color.Gray, 0.5f);
                Pdf.DrawLine(pen, contentRect.Left, contentRect.Top, contentRect.Right, contentRect.Top);
            }
        }

        /// <summary>
        /// 繪製所有頁面的 Footer。
        /// 在所有內容繪製完成後呼叫。
        /// </summary>
        protected virtual void RenderFooters()
        {
            if (!HeaderFooter.ShowFooter)
            {
                return;
            }

            int totalPages = Pdf.Pages.Count;
            int savedPage = Pdf.CurrentPage;

            for (int page = HeaderFooter.StartPage; page < totalPages; page++)
            {
                Pdf.CurrentPage = page;
                RectangleF contentRect = PdfGlobalConfig.GetContentRect(Pdf.PageRectangle);
                RectangleF footerRect = HeaderFooter.GetFooterRect(contentRect);

                // 若有自訂繪製邏輯則使用
                if (HeaderFooter.OnRenderFooter != null)
                {
                    HeaderFooter.OnRenderFooter(Pdf, page, totalPages, footerRect);
                }
                else
                {
                    // 預設繪製邏輯
                    RenderDefaultFooter(page, totalPages, contentRect, footerRect);
                }
            }

            // 還原到原本頁面
            Pdf.CurrentPage = savedPage;
            Logger.Debug($"已繪製 Footer 至 {totalPages - HeaderFooter.StartPage} 頁");
        }

        /// <summary>
        /// 繪製預設 Footer。
        /// </summary>
        protected virtual void RenderDefaultFooter(int pageIndex, int totalPages, RectangleF contentRect, RectangleF footerRect)
        {
            // 繪製頂線
            if (HeaderFooter.DrawFooterLine)
            {
                GcPen pen = new GcPen(Color.Gray, 0.5f);
                Pdf.DrawLine(pen, contentRect.Left, contentRect.Bottom, contentRect.Right, contentRect.Bottom);
                Logger.Info("Footer中頂線位置: 781.89 pt; 頁面總高度 (A4)：841.89 pt; 垂直邊距 (MarginVertical)：預設為 60 pt");
            }

            // 繪製 Footer 文字（左側）
            string footerText = HeaderFooter.FooterText;
            if (string.IsNullOrEmpty(footerText))
            {
                // 預設格式: Created by {SoftwareName} v{Version} (Exported by {Operator} {Time}) {RUO}
                string ruoStr = HeaderFooter.IsResearchUseOnly ? $"       {HeaderFooter.ResearchUseOnlyText}" : "";
                footerText = $"Created by {HeaderFooter.SoftwareNameText} v{HeaderFooter.VersionText} (Exported by {HeaderFooter.OperatorName} {HeaderFooter.ExportedTime}){ruoStr}";
            }

            if (!string.IsNullOrEmpty(footerText))
            {
                C1StringFormat sfLeft = new C1StringFormat
                {
                    Alignment = C1.Util.HorizontalAlignment.Left, // 修正對齊
                    LineAlignment = C1.Util.VerticalAlignment.Center
                };
                Pdf.DrawString(footerText, HeaderFooter.FooterFont, HeaderFooter.FooterColor, footerRect, sfLeft);
            }

            // 繪製頁碼（右側）
            if (HeaderFooter.ShowPageNumber)
            {
                C1StringFormat sfRight = new C1StringFormat
                {
                    Alignment = C1.Util.HorizontalAlignment.Right,
                    LineAlignment = C1.Util.VerticalAlignment.Center
                };
                string pageText = string.Format(HeaderFooter.PageNumberFormat, pageIndex + 1, totalPages);
                Pdf.DrawString(pageText, HeaderFooter.FooterFont, HeaderFooter.FooterColor, footerRect, sfRight);
            }
        }

        /// <summary>
        /// 非同步產生報表。
        /// 流程：InitializeDocument → RenderContent → RenderHeaders → RenderFooters → Save
        /// </summary>
        /// <param name="outputPath">輸出路徑</param>
        public virtual async Task GenerateAsync(string outputPath)
        {
            await Task.Run(() =>
            {
                try
                {
                    Logger.Info($"開始產生報表: {ReportTitle}");

                    // 1. 初始化文件
                    InitializeDocument();

                    // 2. 繪製內容（由子類別實作）
                    RenderContent();

                    // 3. 繪製 Header（迴圈所有頁面）
                    RenderHeaders();

                    // 4. 繪製 Footer（迴圈所有頁面）
                    RenderFooters();

                    // 5. 確保輸出目錄存在
                    string? directory = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    // 6. 儲存檔案
                    Pdf.Save(outputPath);
                    Logger.Info($"報表已儲存至: {outputPath}");
                }
                catch (Exception ex)
                {
                    Logger.Error($"產生報表失敗: {ex.Message}", ex);
                    throw;
                }
            });
        }

        /// <summary>
        /// 繪製報表內容（由子類別實作）。
        /// </summary>
        protected abstract void RenderContent();

        /// <summary>
        /// 在記憶體中初始化 PDF 文件（不輸出檔案）。
        /// 此方法僅準備空白頁面，不繪製任何內容。
        /// 完成後可呼叫 DrawKitInfoTable 等方法新增內容，最後呼叫 ExportToFile 輸出。
        /// </summary>
        public virtual void InitializeInMemory()
        {
            Logger.Info($"開始在記憶體中建立報表: {ReportTitle}");
            
            // 初始化文件
            InitializeDocument();

            // 計算區域矩形
            CalculateAreaRectangles();
            
            // 設定 CurrentRect 為內容區域起始位置
            CurrentRect = ContentAreaRect;
            
            // 注意：此處不呼叫 RenderContent()，讓使用者透過 DrawKitInfoTable 等方法手動新增內容。
            // Header/Footer 將在 ExportToFile 時繪製。
            
            Logger.Info("PDF 已在記憶體中建立完成 (空白頁面)");
        }

        /// <summary>
        /// 計算 Header/Footer/Content 區域矩形。
        /// </summary>
        protected virtual void CalculateAreaRectangles()
        {
            RectangleF rcPage = GetPageRect(Pdf);
            
            // Header 區域 (根據 ContentRect 往上推)
            HeaderAreaRect = HeaderFooter.GetHeaderRect(rcPage);
            
            // Footer 區域 (根據 ContentRect 往下推)
            FooterAreaRect = HeaderFooter.GetFooterRect(rcPage);
            
            // 內容區域
            // 若有 Header，內容起始點應為 Header 底部保留一點間距
            // 若無 Header，內容起始點為 rcPage.Top
            // 但使用者要求 "Current Y 是 第一頁的y值"，通常指向 rcPage.Top (內容區起點)
            // 且 "Header / footer 的區域要保留"，表示不應覆蓋
            
            float contentTop = rcPage.Top;
            if (HeaderFooter.ShowHeader)
            {
               // 確保內容不蓋到 Header (Header 在 Margin 內，通常在 rcPage 上方)
               // 但若 Header 高度較高或 Margin 較小，需檢查重疊
               if (HeaderAreaRect.Bottom > contentTop)
               {
                   contentTop = HeaderAreaRect.Bottom + 5;
               }
            }

            float contentBottom = rcPage.Bottom;
            if (HeaderFooter.ShowFooter)
            {
                if (FooterAreaRect.Top < contentBottom)
                {
                    contentBottom = FooterAreaRect.Top - 5;
                }
            }

            ContentAreaRect = new RectangleF(
                rcPage.X,
                contentTop,
                rcPage.Width,
                contentBottom - contentTop
            );
            
            Logger.Debug($"Header: {HeaderAreaRect}, Footer: {FooterAreaRect}, Content: {ContentAreaRect}, CurrentY: {contentTop}");
        }

        /// <summary>
        /// 取得頁面內容區域 (類似 rcPage.Inflate(-30, -60))。
        /// </summary>
        protected RectangleF GetPageRect(C1PdfDocument pdf)
        {
            // 使用全域設定取得內容區域 (已套用 Margin)
            return PdfGlobalConfig.GetContentRect(pdf.PageRectangle);
        }

        /// <summary>
        /// 將記憶體中的 PDF 輸出為檔案。
        /// 必須先呼叫 InitializeInMemory 或 GenerateAsync。
        /// </summary>
        /// <param name="outputPath">輸出檔案路徑</param>
        public virtual void ExportToFile(string outputPath)
        {
            if (!IsPdfInitialized)
            {
                throw new InvalidOperationException("PDF 尚未初始化，請先呼叫 InitializeInMemory()");
            }

            string? directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // 輸出前繪製 Header/Footer
            RenderHeaders();
            RenderFooters();

            Pdf.Save(outputPath);
            _isExported = true;
            Logger.Info($"PDF 已輸出至: {outputPath}");
        }

        /// <summary>
        /// 繪製 Kit Info 表格至目前 Y 位置。
        /// </summary>
        /// <param name="data">Key-Value 資料字典</param>
        /// <returns>繪製區域的矩形</returns>
        public virtual RectangleF DrawKitInfoTable(Dictionary<string, string> data)
        {
            if (!IsPdfInitialized)
            {
                throw new InvalidOperationException("PDF 尚未初始化，請先呼叫 InitializeInMemory()");
            }

            if (data == null || data.Count == 0)
            {
                Logger.Info("Kit Info 資料為空，跳過繪製");
                return RectangleF.Empty;
            }

            Logger.Info($"開始繪製 Kit Info 表格，共 {data.Count} 筆資料");

            float rowHeight = 15f;
            float labelWidth = 120f;
            float valueWidth = CurrentRect.Width - labelWidth;
            float startY = CurrentRect.Y;

            C1StringFormat sfLabel = new C1StringFormat
            {
                Alignment = C1.Util.HorizontalAlignment.Left,
                LineAlignment = C1.Util.VerticalAlignment.Center
            };

            C1StringFormat sfValue = new C1StringFormat
            {
                Alignment = C1.Util.HorizontalAlignment.Left,
                LineAlignment = C1.Util.VerticalAlignment.Center
            };

            float currentY = startY;
            bool pageBreakHappened = false;

            foreach (var kvp in data)
            {
                // 檢查是否需要換頁
                if (currentY + rowHeight > ContentAreaRect.Bottom)
                {
                    AddNewPage();
                    CalculateAreaRectangles();
                    currentY = ContentAreaRect.Y;
                    pageBreakHappened = true;
                }

                // 繪製 Label
                RectangleF labelRect = new RectangleF(ContentAreaRect.X, currentY, labelWidth, rowHeight);
                Pdf.DrawString(kvp.Key + ":", NoteFont, Color.Black, labelRect, sfLabel);

                // 繪製 Value
                RectangleF valueRect = new RectangleF(ContentAreaRect.X + labelWidth, currentY, valueWidth, rowHeight);
                Pdf.DrawString(kvp.Value, NoteFont, Color.Black, valueRect, sfValue);

                currentY += rowHeight;
            }

            // 繪製表格底線
            Pdf.DrawLine(new GcPen(Color.Gray, 0.2f), PageRect.Left, currentY, PageRect.Left + CurrentRect.Width, currentY);

            // 更新 CurrentRect
            CurrentRect = new RectangleF(
                CurrentRect.X,
                currentY,
                CurrentRect.Width,
                ContentAreaRect.Bottom - currentY
            );

            // 計算回傳區域
            // 若發生換頁，簡單回傳最後一頁的繪製區域 (使用者通常關心最後狀態，或是這只是一個單頁表格)
            // 若需完整跨頁區域可能需回傳 List<RectangleF>，目前先以上層需求為主。
            float resultY = pageBreakHappened ? ContentAreaRect.Y : startY;
            float resultHeight = currentY - resultY;

            RectangleF usedRect = new RectangleF(ContentAreaRect.X, resultY, ContentAreaRect.Width, resultHeight);
            Logger.Info($"Kit Info 表格繪製完成，CurrentY: {CurrentY}, Used Rect: {usedRect}");
            return usedRect;
        }

        /// <summary>
        /// 繪製水平分隔線。
        /// </summary>
        /// <param name="lineColor">線條顏色 (預設: Gray)</param>
        /// <param name="lineWidth">線條寬度 (預設: 0.2f)</param>
        /// <param name="addSpacingAfter">繪製後增加間距 (預設: 0)</param>
        /// <param name="x">起始 X 位置 (null 則使用內容區左側)</param>
        /// <param name="length">線條長度 (null 則使用內容區寬度)</param>
        /// <returns>繪製後的 CurrentY 位置</returns>
        public virtual float DrawHorizontalLine(
            Color? lineColor = null,
            float lineWidth = 0.2f,
            float addSpacingAfter = 0f,
            float? x = null,
            float? length = null)
        {
            if (!IsPdfInitialized)
            {
                throw new InvalidOperationException("PDF 尚未初始化或已匯出封裝，請先呼叫 InitializeInMemory()");
            }

            Color color = lineColor ?? Color.Gray;
            float startX = x ?? PageRect.Left;
            float lineLen = length ?? PageRect.Width;
            
            // 繪製水平線
            PointF startPoint = new PointF(startX, CurrentY);
            PointF endPoint = new PointF(startX + lineLen, CurrentY);
            
            Pdf.DrawLine(new GcPen(color, lineWidth), startPoint, endPoint);
            
            Logger.Info($"繪製水平線: Y={CurrentY}, X={startX}, Length={lineLen}, 寬度={lineWidth}, 顏色={color.Name}");

            // 更新 CurrentRect (增加間距)
            if (addSpacingAfter > 0)
            {
                CurrentRect = new RectangleF(
                    CurrentRect.X,
                    CurrentRect.Y + addSpacingAfter,
                    CurrentRect.Width,
                    CurrentRect.Height - addSpacingAfter
                );
            }

            return CurrentY;
        }

        /// <summary>
        /// 強制換頁。
        /// </summary>
        public virtual void PageBreak()
        {
            if (!IsPdfInitialized)
            {
                throw new InvalidOperationException("PDF 尚未初始化，無法換頁。");
            }

            AddNewPage();
            // AddNewPage 內部已經會重置 CurrentRect 為新頁面的 PageRect
            Logger.Info($"執行強制換頁: 新頁面索引 {CurrentPageIndex}");
        }

        /// <summary>
        /// 繪製簽名區 (Operator / Director)。
        /// </summary>
        public void DrawSignatureArea()
        {
            float requiredHeight = 120; // 預留高度
            
            // 檢查換頁
            if (CurrentY > PageRect.Bottom - requiredHeight)
            {
                AddNewPage();
                // 換頁後增加頂部間距 (參考程式碼: rc.Y += 30)
                CurrentRect = new RectangleF(CurrentRect.X, PageRect.Top + 30, CurrentRect.Width, PageRect.Height - 30);
            }

            C1Font font = new C1Font("Arial", 9, C1.Util.FontStyle.Bold);

            // 1. Operator
            // 參考程式碼: size.Height += 6
            SizeF sizeOp = Pdf.MeasureString("Operator", font, PageRect.Width);
            float hOp = sizeOp.Height + 6;
            
            RectangleF rcOp = new RectangleF(PageRect.Left, CurrentY, PageRect.Width, hOp);
            Pdf.DrawString("Operator", font, Color.Black, rcOp);
            
            // 下移: 文字高度 + 30 (Gap)
            float moveY = hOp + 30;
            CurrentRect = new RectangleF(CurrentRect.X, CurrentRect.Y + moveY, CurrentRect.Width, CurrentRect.Height - moveY);

            // Line 1 (寬度為頁面的一半)
            // DrawHorizontalLine 會自動更新 CurrentY (且預設無額外間距，我們需要自己控制)
            // 參考: 畫線後 rc.Y += 15
            DrawHorizontalLine(Color.Gray, 0.2f, 15, PageRect.Left, PageRect.Width / 2);

            // 2. Director
            // 參考程式碼: size.Height += 6
            SizeF sizeDir = Pdf.MeasureString("Director", font, PageRect.Width);
            float hDir = sizeDir.Height + 6;
            
            RectangleF rcDir = new RectangleF(PageRect.Left, CurrentY, PageRect.Width, hDir);
            Pdf.DrawString("Director", font, Color.Black, rcDir);

            // 下移: 文字高度 + 30 (Gap)
            moveY = hDir + 30;
            CurrentRect = new RectangleF(CurrentRect.X, CurrentRect.Y + moveY, CurrentRect.Width, CurrentRect.Height - moveY);

            // Line 2
            // 參考: 畫線後 rc.Y += 15
            DrawHorizontalLine(Color.Gray, 0.2f, 15, PageRect.Left, PageRect.Width / 2);
            
            Logger.Info("已繪製簽名區");
        }

        /// <summary>
        /// 插入垂直間隔區域。若超出邊界則自動換頁。
        /// </summary>
        /// <param name="height">間隔高度 (pt)</param>
        public virtual void AddVerticalSpacing(float height)
        {
            if (!IsPdfInitialized)
            {
                throw new InvalidOperationException("PDF 尚未初始化，無法插入間隔。");
            }

            float targetY = CurrentY + height;

            if (targetY > PageRect.Bottom)
            {
                Logger.Info($"間隔後位置 ({targetY}) 超出頁面底部 ({PageRect.Bottom})，執行自動換頁。");
                PageBreak();
                // 換頁後，從新頁面頂部開始，可選擇是否仍要插入間隔 (通常換頁後不需要額外頂部間隔，或可視需求調整)
                // 若需在新頁面保留間距，可在此處遞迴呼叫或直接更新 Rect，這裡暫時策略為「換頁即視為已分隔」
            }
            else
            {
                // 正常更新 CurrentRect
                CurrentRect = new RectangleF(
                    CurrentRect.X,
                    CurrentRect.Y + height,
                    CurrentRect.Width,
                    CurrentRect.Height - height
                );
                Logger.Info($"插入垂直間隔: {height} pt, 新 CurrentY: {CurrentY}");
            }
        }

        /// <summary>
        /// 使用 TbSetting 繪製 Kit Info 表格 (V1 版本)。
        /// </summary>
        /// <param name="data">4 欄 DataTable 資料</param>
        /// <param name="renderMethod">繪製方法名稱 (預設: RenderTable_V1)</param>
        /// <param name="styleName">樣式名稱 (預設: TbSetting1)</param>
        /// <returns>繪製區域的矩形</returns>
        public virtual RectangleF DrawKitInfoTableWithStyle(
            System.Data.DataTable data,
            string renderMethod = "RenderTable_V1",
            string styleName = "TbSetting1")
        {
            if (!IsPdfInitialized)
            {
                throw new InvalidOperationException("PDF 尚未初始化，請先呼叫 InitializeInMemory()");
            }

            if (data == null || data.Rows.Count == 0)
            {
                Logger.Info("Kit Info 資料為空，跳過繪製");
                return RectangleF.Empty;
            }

            Logger.Info($"開始繪製 Kit Info 表格，方法: {renderMethod}, 樣式: {styleName}, 共 {data.Rows.Count} 列");

            // 取得樣式設定
            TbSetting setting = GetTbSetting(styleName);

            // 根據方法名稱選擇繪製方式
            RectangleF result;
            switch (renderMethod)
            {
                case "RenderTable_V1":
                default:
                    result = RenderTable_V1(data, setting);
                    break;
            }

            Logger.Info($"Kit Info 表格繪製完成，CurrentY: {CurrentY}, Used Rect: {result}");
            return result;
        }

        /// <summary>
        /// 取得表格樣式設定。
        /// </summary>
        /// <param name="styleName">樣式名稱</param>
        /// <returns>TbSetting 實例</returns>
        protected virtual TbSetting GetTbSetting(string styleName)
        {
            switch (styleName)
            {
                case "TbSetting1":
                default:
                    return TbSetting.CreateTbSetting1(NoteFont);
            }
        }

        /// <summary>
        /// RenderTable V1 (簡化版)：繪製 DataTable 至 PDF。
        /// </summary>
        /// <param name="dt">資料表</param>
        /// <param name="setting">表格樣式設定</param>
        /// <returns>繪製區域的矩形</returns>
        public virtual RectangleF RenderTable_V1(System.Data.DataTable dt, TbSetting setting)
        {
            if (dt == null || dt.Rows.Count == 0)
            {
                return RectangleF.Empty;
            }

            float startY = CurrentRect.Y;
            float currentY = startY;
            float startX = ContentAreaRect.X;
            float tableWidth = ContentAreaRect.Width;

            // 計算欄位寬度
            float totalWeight = 0f;
            foreach (System.Data.DataColumn col in dt.Columns)
            {
                if (setting.ColumnWidthFactors.TryGetValue(col.ColumnName, out float weight))
                {
                    totalWeight += weight;
                }
                else
                {
                    totalWeight += 1f; // 預設權重
                }
            }

            float[] columnWidths = new float[dt.Columns.Count];
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                string colName = dt.Columns[i].ColumnName;
                float weight = setting.ColumnWidthFactors.TryGetValue(colName, out float w) ? w : 1f;
                columnWidths[i] = tableWidth * (weight / totalWeight);
            }

            // 設定字型與格式 (使用 C1Font)
            C1Font rowFont = setting.RowFont ?? NoteFont;

            // 建立 C1StringFormat
            C1StringFormat c1Sf = new C1StringFormat
            {
                Alignment = C1.Util.HorizontalAlignment.Left,
                LineAlignment = C1.Util.VerticalAlignment.Center
            };

            // 套用 TbSetting 的設定
            if (setting.RowStringFormat != null)
            {
                c1Sf.Alignment = setting.RowStringFormat.Alignment == System.Drawing.StringAlignment.Near ? C1.Util.HorizontalAlignment.Left :
                                 setting.RowStringFormat.Alignment == System.Drawing.StringAlignment.Center ? C1.Util.HorizontalAlignment.Center :
                                 C1.Util.HorizontalAlignment.Right;
                c1Sf.LineAlignment = setting.RowStringFormat.LineAlignment == System.Drawing.StringAlignment.Near ? C1.Util.VerticalAlignment.Top :
                                     setting.RowStringFormat.LineAlignment == System.Drawing.StringAlignment.Center ? C1.Util.VerticalAlignment.Center :
                                     C1.Util.VerticalAlignment.Bottom;
            }

            float rowHeight = rowFont.Size + 8f + setting.RowHeightMargin;
            bool pageBreakHappened = false;

            // 繪製資料列
            foreach (System.Data.DataRow dr in dt.Rows)
            {
                // 檢查換頁
                if (currentY + rowHeight > ContentAreaRect.Bottom)
                {
                    AddNewPage();
                    CalculateAreaRectangles();
                    currentY = ContentAreaRect.Y;
                    startX = ContentAreaRect.X;
                    pageBreakHappened = true;
                }

                float cellX = startX;

                for (int colIdx = 0; colIdx < dt.Columns.Count; colIdx++)
                {
                    string text = dr[colIdx]?.ToString() ?? string.Empty;
                    float cellWidth = columnWidths[colIdx];

                    // 檢查欄位合併 (只有當後續欄位為空時才合併)
                    if (setting.MergeColumns.TryGetValue(colIdx, out int mergeCount) && mergeCount > 1)
                    {
                        // 檢查後續欄位是否為空，若為空則進行合併
                        bool shouldMerge = true;
                        for (int m = 1; m < mergeCount && (colIdx + m) < dt.Columns.Count; m++)
                        {
                            string nextCellText = dr[colIdx + m]?.ToString() ?? string.Empty;
                            if (!string.IsNullOrEmpty(nextCellText))
                            {
                                shouldMerge = false;
                                break;
                            }
                        }

                        if (shouldMerge)
                        {
                            // 合併多個欄位的寬度
                            for (int m = 1; m < mergeCount && (colIdx + m) < dt.Columns.Count; m++)
                            {
                                cellWidth += columnWidths[colIdx + m];
                            }
                            // 繪製合併儲存格
                            RectangleF mergedRect = new RectangleF(cellX, currentY, cellWidth, rowHeight);
                            if (setting.RowType == TbSetting.RowStyle.HasRowBorder)
                            {
                                Pdf.DrawRectangle(new GcPen(Color.Gray, 0.5f), mergedRect);
                            }
                            RectangleF mergedTextRect = new RectangleF(cellX + 2, currentY, cellWidth - 4, rowHeight);
                            Pdf.DrawString(text, rowFont, Color.Black, mergedTextRect, c1Sf);
                            cellX += cellWidth;
                            colIdx += mergeCount - 1; // 跳過已合併的欄位
                            continue;
                        }
                    }

                    RectangleF cellRect = new RectangleF(cellX, currentY, cellWidth, rowHeight);

                    // 繪製框線 (若 RowStyle 為 HasRowBorder)
                    if (setting.RowType == TbSetting.RowStyle.HasRowBorder)
                    {
                        Pdf.DrawRectangle(new GcPen(Color.Gray, 0.5f), cellRect);
                    }

                    // 繪製文字 (加入內縮)
                    RectangleF textRect = new RectangleF(cellX + 2, currentY, cellWidth - 4, rowHeight);
                    Pdf.DrawString(text, rowFont, Color.Black, textRect, c1Sf);

                    cellX += cellWidth;
                }

                currentY += rowHeight;
            }

            // 繪製表格底線
            Pdf.DrawLine(new GcPen(Color.Gray, 0.2f), startX, currentY, startX + tableWidth, currentY);

            // 更新 CurrentRect
            CurrentRect = new RectangleF(
                CurrentRect.X,
                currentY,
                CurrentRect.Width,
                ContentAreaRect.Bottom - currentY
            );

            // 計算回傳區域
            float resultY = pageBreakHappened ? ContentAreaRect.Y : startY;
            float resultHeight = currentY - resultY;

            return new RectangleF(ContentAreaRect.X, resultY, tableWidth, resultHeight);
        }

        protected RectangleF RenderPcncNote(C1PdfDocument _c1pdf, IEnumerable<string> noteLines, RectangleF rcPage, RectangleF rc, float fontSize = 7)
        {
            // 設定格式
            C1StringFormat sfRow = new C1StringFormat();
            sfRow.LineAlignment = C1.Util.VerticalAlignment.Center;
            sfRow.Alignment = C1.Util.HorizontalAlignment.Left; // Near = Left

            C1Font font = new C1Font("Arial", fontSize, C1.Util.FontStyle.Regular);
            
            foreach (string line in noteLines)
            {
                // 計算每一行的高度
                SizeF size = _c1pdf.MeasureString(line, font, rcPage.Width / 2);
                size.Height += 6;

                // 檢查是否需要換頁
                if (rc.Y + size.Height > rcPage.Bottom)
                {
                    _c1pdf.NewPage();
                    rc.Y = rcPage.Top + 20; // 換頁後預留頂部邊距
                }

                // 設定繪製區域
                rc.Width = rcPage.Width / 2;
                rc.Height = size.Height;
                rc.X = rcPage.Left;

                // 繪製文字
                rc.Inflate(-2, 0); // 內縮
                _c1pdf.DrawString(line, font, Color.Black, rc, sfRow);
                rc.Inflate(2, 0); // 還原

                // 移動到下一行位置
                rc.Offset(0, rc.Height);
            }

            return rc;
        }

        protected RectangleF RenderPcncTable(C1PdfDocument _c1pdf, List<PcncTableItem> items, RectangleF rcPage, RectangleF rc, float fontSize = 8)
        {
            if (items == null || items.Count == 0) return rc;

            // 定義欄位標題 - 配合 CSV 只有 4 欄
            string[] headers = { "Well ID", "Control", "Result", "Flag" };
            
            // 欄位寬度設定 (25%, 25%, 25%, 25%)
            float[] colWidths = { 
                rcPage.Width * 0.25f, 
                rcPage.Width * 0.25f, 
                rcPage.Width * 0.25f, 
                rcPage.Width * 0.25f 
            };
            
            // 設定並累積 X 座標
            float[] colX = new float[colWidths.Length];
            colX[0] = rcPage.Left;
            for (int i = 1; i < colWidths.Length; i++)
            {
                colX[i] = colX[i-1] + colWidths[i-1];
            }

            // 字型與格式
            C1Font fontHeader = new C1Font("Arial", fontSize, C1.Util.FontStyle.Bold);
            C1Font fontItem = new C1Font("Arial", fontSize, C1.Util.FontStyle.Regular);
            
            // 對齊方式
            C1StringFormat sfLeft = new C1StringFormat();
            sfLeft.LineAlignment = C1.Util.VerticalAlignment.Center;
            sfLeft.Alignment = C1.Util.HorizontalAlignment.Left;
            
            // 計算標題高度
            float headerHeight = 0;
            for (int i = 0; i < headers.Length; i++)
            {
                float h = _c1pdf.MeasureString(headers[i], fontHeader, colWidths[i]).Height + 6;
                if (h > headerHeight) headerHeight = h;
            }

            // 檢查換頁 (標題)
            if (rc.Y + headerHeight > rcPage.Bottom)
            {
                _c1pdf.NewPage();
                rc.Y = rcPage.Top + 20;
            }
            
            float tableTopY = rc.Y;

            // 繪製標題文字
            for (int i = 0; i < headers.Length; i++)
            {
                RectangleF rcCell = new RectangleF(colX[i], rc.Y, colWidths[i], headerHeight);
                rcCell.Inflate(-2, 0);
                _c1pdf.DrawString(headers[i], fontHeader, Color.Black, rcCell, sfLeft);
            }
            
            // 畫標題分隔線 (Header Bottom / Grid Top)
            // 參考程式碼: Header 上方有線, 下方有線. 
            // 這裡實作: Table Top Line (Header Top) AND Header Bottom Line.
            _c1pdf.DrawLine(new GcPen(Color.Gray, 1), rcPage.Left, rc.Y, rcPage.Right, rc.Y); // Top
            rc.Y += headerHeight;
            _c1pdf.DrawLine(new GcPen(Color.Gray, 1), rcPage.Left, rc.Y, rcPage.Right, rc.Y); // Header Bottom

            // 繪製資料列
            foreach (var item in items)
            {
                string[] values = { item.WellId, item.Control, item.Result, item.Flag };
                
                // 判斷是否為紅色 (Fail)
                bool isRed = item.Result.Trim().Equals("Fail", StringComparison.OrdinalIgnoreCase);

                // 計算列高
                float rowHeight = 0;
                for (int i = 0; i < values.Length; i++)
                {
                    float h = _c1pdf.MeasureString(values[i], fontItem, colWidths[i] - 4).Height + 6;
                    if (h > rowHeight) rowHeight = h;
                }

                // 檢查換頁
                if (rc.Y + rowHeight > rcPage.Bottom)
                {
                    // 換頁前, 繪製當前頁的 Table Bottom Line (若需要封閉當前頁表格)
                    _c1pdf.DrawLine(new GcPen(Color.Gray, 1), rcPage.Left, rc.Y, rcPage.Right, rc.Y);
                    
                    // 繪製當前頁的左右邊界 (從 tableTopY 到 rc.Y)
                    _c1pdf.DrawLine(new GcPen(Color.Gray, 1), rcPage.Left, tableTopY, rcPage.Left, rc.Y);
                    _c1pdf.DrawLine(new GcPen(Color.Gray, 1), rcPage.Right, tableTopY, rcPage.Right, rc.Y);

                    _c1pdf.NewPage();
                    rc.Y = rcPage.Top + 20;
                    tableTopY = rc.Y; // 重置新頁面的 tableTopY
                    
                    // 換頁後, 補上頂線
                    _c1pdf.DrawLine(new GcPen(Color.Gray, 1), rcPage.Left, rc.Y, rcPage.Right, rc.Y);
                }

                // 繪製當前列內容
                for (int i = 0; i < values.Length; i++)
                {
                    RectangleF rcCell = new RectangleF(colX[i], rc.Y, colWidths[i], rowHeight);
                    
                    // 內容
                    rcCell.Inflate(-2, 0);
                    
                    // 根據欄位決定顏色: 只有 Result (index 2) 且為 Fail 時才紅
                    Color cellColor = Color.Black;
                    if (i == 2 && isRed)
                    {
                        cellColor = Color.Red;
                    }

                    _c1pdf.DrawString(values[i], fontItem, cellColor, rcCell, sfLeft);
                }
                
                // 不畫每列底線 (參考風格: 內部無線條)
                rc.Y += rowHeight;
            }

            // 表格結束, 繪製 Bottom Line
            _c1pdf.DrawLine(new GcPen(Color.Gray, 1), rcPage.Left, rc.Y, rcPage.Right, rc.Y);
            
            // 繪製左右邊界 (從 tableTopY 到 rc.Y)
            _c1pdf.DrawLine(new GcPen(Color.Gray, 1), rcPage.Left, tableTopY, rcPage.Left, rc.Y);
            _c1pdf.DrawLine(new GcPen(Color.Gray, 1), rcPage.Right, tableTopY, rcPage.Right, rc.Y);

            // 更新 rc
            rc.Height = 0; 
            return rc;
        }

        /// <summary>
        /// 繪製 PC/NC Fail Detail Table (含自動分組合併、奇偶列變色、per-cell 顏色覆寫、換頁處理)。
        /// 依據 DetailTableStyle 控制所有表格樣式，包含透過 style.PageBottomMargin 控制換頁時距離 Footer 的預留緩衝區 (預設繼承自 PdfGlobalConfig.DefaultPageBottomMargin)。
        /// </summary>
        /// <param name="_c1pdf">C1PdfDocument 實例</param>
        /// <param name="items">扁平資料列表 (已排序)</param>
        /// <param name="style">表格樣式定義</param>
        /// <param name="rcPage">頁面可用區域</param>
        /// <param name="rc">當前繪製位置</param>
        /// <returns>更新後的繪製位置</returns>
        protected RectangleF RenderPcncDetailTable(C1PdfDocument _c1pdf, List<PcncDetailItem> items, DetailTableStyle style, RectangleF rcPage, RectangleF rc)
        {
            if (items == null || items.Count == 0) return rc;

            // 驗證樣式參數
            style.Validate();

            int numOfCol = style.Headers.Length;
            float separatorGap = Math.Max(style.BorderWidth, 1.0f); // 定義分隔線間隙，避免背景色覆蓋線條

            // === 計算欄寬與 X 座標 ===
            float baseUnit = rcPage.Width / style.WidthBaseDivider;
            float[] colWidths = new float[numOfCol];
            for (int i = 0; i < numOfCol; i++)
            {
                colWidths[i] = baseUnit * style.ColumnWidthFactors[i] + style.ColumnWidthOffsets[i];
            }

            float[] colX = new float[numOfCol];
            colX[0] = rcPage.Left;
            for (int i = 1; i < numOfCol; i++)
            {
                colX[i] = colX[i - 1] + colWidths[i - 1];
            }

            // === 字型與格式 ===
            C1Font fontHeader = new C1Font("Arial", style.FontSize, C1.Util.FontStyle.Bold);
            C1Font fontItem = new C1Font("Arial", style.FontSize, C1.Util.FontStyle.Regular);

            C1StringFormat sfLeft = new C1StringFormat();
            sfLeft.LineAlignment = C1.Util.VerticalAlignment.Center;
            sfLeft.Alignment = C1.Util.HorizontalAlignment.Left;

            C1StringFormat sfRight = new C1StringFormat();
            sfRight.LineAlignment = C1.Util.VerticalAlignment.Center;
            sfRight.Alignment = C1.Util.HorizontalAlignment.Right;

            C1StringFormat sfCenter = new C1StringFormat();
            sfCenter.LineAlignment = C1.Util.VerticalAlignment.Center;
            sfCenter.Alignment = C1.Util.HorizontalAlignment.Center;

            // 對齊方式查找
            C1StringFormat GetAlignment(int colIdx)
            {
                int align = style.ColumnAlignments[colIdx];
                return align switch
                {
                    1 => sfRight,
                    2 => sfCenter,
                    _ => sfLeft
                };
            }

            // === 可繪製底部邊界 ===
            float drawableBottom = rcPage.Bottom
                - style.PageBottomMargin
                - (style.HasFooterNote ? style.FooterNoteHeight : 0);

            // === 計算 Header 高度 ===
            float headerHeight = 0;
            for (int i = 0; i < numOfCol; i++)
            {
                float h = _c1pdf.MeasureString(style.Headers[i], fontHeader, colWidths[i] - style.CellPadding * 2).Height + 6;
                if (h > headerHeight) headerHeight = h;
            }

            // === 自動分組 (合併邏輯) ===
            List<List<PcncDetailItem>> groups;
            if (style.EnableColumnMerge && style.MergeColumnIndices.Length > 0)
            {
                groups = BuildMergeGroups(items, style.MergeColumnIndices);
            }
            else
            {
                // 不合併: 每列各自一組
                groups = new List<List<PcncDetailItem>>();
                foreach (var item in items)
                {
                    groups.Add(new List<PcncDetailItem> { item });
                }
            }

            // === 找到列分隔線起始欄 ===
            int firstNonMergeCol = 0;
            if (style.EnableColumnMerge && style.RowSeparator == RowSeparatorMode.SkipMergedColumns)
            {
                for (int i = 0; i < numOfCol; i++)
                {
                    if (!Array.Exists(style.MergeColumnIndices, idx => idx == i))
                    {
                        firstNonMergeCol = i;
                        break;
                    }
                }
            }

            // === Helper: Normalize Text (unescape \\r\\n and add ZWS for wrapping) ===
            string NormalizeText(string text)
            {
                if (string.IsNullOrEmpty(text)) return text;
                string s = text.Replace("\\r\\n", "\n").Replace("\\n", "\n").Replace("\\r", "\n");
                
                // Add Zero Width Space (\u200B) after characters to force "Break Anywhere" behavior
                // 這對 DNA 序列或長字串非常重要，避免因標點符號而過早換行
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                foreach (char c in s)
                {
                    sb.Append(c);
                    if (c != '\n') sb.Append('\u200B');
                }
                return sb.ToString();
            }

            // === Helper: 繪製 Header ===
            void DrawHeader()
            {
                // 頂線
                _c1pdf.DrawLine(new GcPen(style.BorderColor, style.BorderWidth), rcPage.Left, rc.Y, rcPage.Right, rc.Y);

                for (int i = 0; i < numOfCol; i++)
                {
                    RectangleF rcCell = new RectangleF(colX[i], rc.Y, colWidths[i], headerHeight);
                    rcCell.Inflate(-style.CellPadding, 0);
                    _c1pdf.DrawString(style.Headers[i], fontHeader, Color.Black, rcCell, GetAlignment(i));
                }

                rc.Y += headerHeight;
                // 底線
                _c1pdf.DrawLine(new GcPen(style.BorderColor, style.BorderWidth), rcPage.Left, rc.Y, rcPage.Right, rc.Y);
                rc.Y += separatorGap; // 使用統一的間隙變數
            }

            // === Helper: 繪製外框 Left + Right ===
            void DrawOuterBorders(float topY, float bottomY)
            {
                _c1pdf.DrawLine(new GcPen(style.BorderColor, style.BorderWidth), rcPage.Left, topY, rcPage.Left, bottomY);
                _c1pdf.DrawLine(new GcPen(style.BorderColor, style.BorderWidth), rcPage.Right, topY, rcPage.Right, bottomY);
            }

            // === Helper: 繪製合併欄 ===
            void DrawMergeColumns(PcncDetailItem firstItem, float topY, float height)
            {
                if (!style.EnableColumnMerge) return;

                foreach (int mergeCol in style.MergeColumnIndices)
                {
                    RectangleF rcMerge = new RectangleF(colX[mergeCol], topY, colWidths[mergeCol], height);
                    rcMerge.Inflate(-style.CellPadding, 0);

                    Color textColor = Color.Black;
                    if (firstItem.CellColorOverrides.TryGetValue(mergeCol, out Color overrideColor))
                    {
                        textColor = overrideColor;
                    }

                    _c1pdf.DrawString(NormalizeText(firstItem.GetValueByIndex(mergeCol)), fontItem, textColor, rcMerge, GetAlignment(mergeCol));
                }
            }

            // === Helper: 繪製列分隔線 ===
            void DrawRowSeparator(float y)
            {
                switch (style.RowSeparator)
                {
                    case RowSeparatorMode.FullWidth:
                        _c1pdf.DrawLine(new GcPen(style.BorderColor, style.BorderWidth), rcPage.Left, y, rcPage.Right, y);
                        break;
                    case RowSeparatorMode.SkipMergedColumns:
                        _c1pdf.DrawLine(new GcPen(style.BorderColor, style.BorderWidth), colX[firstNonMergeCol], y, rcPage.Right, y);
                        break;
                    case RowSeparatorMode.None:
                        break;
                }
            }

            // === 開始繪製 ===

            // 檢查換頁 (Header)
            float firstRowHeightEstimate = 14f;
            if (groups.Count > 0 && groups[0].Count > 0)
            {
                PcncDetailItem firstItem = groups[0][0];
                float minH = _c1pdf.MeasureString("Tg", fontItem, 100).Height + 6;
                firstRowHeightEstimate = minH;
                for (int colIdx = 0; colIdx < numOfCol; colIdx++)
                {
                    if (style.EnableColumnMerge && Array.Exists(style.MergeColumnIndices, idx => idx == colIdx)) continue;
                    string cellText = NormalizeText(firstItem.GetValueByIndex(colIdx));
                    if (!string.IsNullOrEmpty(cellText))
                    {
                        float h = _c1pdf.MeasureString(cellText, fontItem, colWidths[colIdx] - style.CellPadding * 2, GetAlignment(colIdx)).Height + 6;
                        if (h > firstRowHeightEstimate) firstRowHeightEstimate = h;
                    }
                }
            }

            if (rc.Y > rcPage.Top + 10f && rc.Y + headerHeight + firstRowHeightEstimate > drawableBottom)
            {
                _c1pdf.NewPage();
                rc.Y = rcPage.Top + style.PageTopMargin;
            }
            else if (rc.Y + headerHeight > drawableBottom)
            {
                _c1pdf.NewPage();
                rc.Y = rcPage.Top + style.PageTopMargin;
            }

            // 強制第一頁的起始位置至少要符合 PageTopMargin (確保與換頁後的視覺一致)
            if (rc.Y < rcPage.Top + style.PageTopMargin)
            {
                rc.Y = rcPage.Top + style.PageTopMargin;
            }

            float tableTopY = rc.Y;
            DrawHeader();

            int dataRowIndex = 0; // 全域列計數器 (跨頁不重置)

            for (int gIdx = 0; gIdx < groups.Count; gIdx++)
            {
                List<PcncDetailItem> group = groups[gIdx];
                float groupTopY = rc.Y;
                float passHeight = 0;

                for (int rIdx = 0; rIdx < group.Count; rIdx++)
                {
                    PcncDetailItem item = group[rIdx];


                    // 計算列高 (僅非合併欄)
                    float minRowHeight = _c1pdf.MeasureString("Tg", fontItem, 100).Height + 6;
                    float rowHeight = minRowHeight;
                    
                    for (int colIdx = 0; colIdx < numOfCol; colIdx++)
                    {
                        if (style.EnableColumnMerge && Array.Exists(style.MergeColumnIndices, idx => idx == colIdx))
                        {
                            continue; // 合併欄不參與逐列高度計算
                        }

                        string cellText = NormalizeText(item.GetValueByIndex(colIdx));
                        if (!string.IsNullOrEmpty(cellText))
                        {
                            float h = _c1pdf.MeasureString(cellText, fontItem, colWidths[colIdx] - style.CellPadding * 2, GetAlignment(colIdx)).Height + 6;
                            if (h > rowHeight) rowHeight = h;
                        }
                    }

                    // 換頁判斷
                    if (rc.Y + rowHeight > drawableBottom)
                    {
                        // 1. 完成當前頁合併欄
                        if (passHeight > 0)
                        {
                            DrawMergeColumns(group[0], groupTopY, passHeight);
                        }

                        // 2. 外框封閉 (需扣除前一個元素累積的間隙，避免雙線)
                        float breakBottomY = rc.Y - separatorGap;
                        Logger.Info($"[PageBreak] TableBottom: {breakBottomY:F2}, FooterTop(Limit): {drawableBottom:F2}, Diff: {(drawableBottom - breakBottomY):F2}, H: {(rcPage.Height + PdfGlobalConfig.MarginVertical * 2):F1}, FooterH: {HeaderFooter.FooterHeight:F1}, FooterBottom: {HeaderFooter.GetFooterRect(rcPage).Bottom:F2}");

                        _c1pdf.DrawLine(new GcPen(style.BorderColor, style.BorderWidth), rcPage.Left, breakBottomY, rcPage.Right, breakBottomY); // Bottom
                        DrawOuterBorders(tableTopY, breakBottomY);

                        // 3. 換頁
                        _c1pdf.NewPage();
                        rc.Y = rcPage.Top + style.PageTopMargin;
                        tableTopY = rc.Y;

                        // 4. 重繪 Header
                        if (style.RedrawHeaderOnNewPage)
                        {
                            DrawHeader();
                        }

                        groupTopY = rc.Y;
                        passHeight = 0;
                        // dataRowIndex 不重置
                    }

                    // 繪製背景色 (奇偶列)
                    if (style.AlternatingRowBackground)
                    {
                        Color bgColor = (dataRowIndex % 2 == 0) ? style.EvenRowColor : style.OddRowColor;

                        // 填充非合併欄的背景
                        if (style.EnableColumnMerge)
                        {
                            float bgStartX = colX[firstNonMergeCol];
                            float bgWidth = rcPage.Right - bgStartX;
                            _c1pdf.FillRectangle(bgColor, new RectangleF(bgStartX, rc.Y, bgWidth, rowHeight));
                        }
                        else
                        {
                            _c1pdf.FillRectangle(bgColor, new RectangleF(rcPage.Left, rc.Y, rcPage.Width, rowHeight));
                        }
                    }

                    // 繪製非合併欄文字
                    for (int colIdx = 0; colIdx < numOfCol; colIdx++)
                    {
                        if (style.EnableColumnMerge && Array.Exists(style.MergeColumnIndices, idx => idx == colIdx))
                        {
                            continue; // 合併欄稍後統一繪製
                        }

                        string cellText = NormalizeText(item.GetValueByIndex(colIdx));
                        RectangleF rcCell = new RectangleF(colX[colIdx], rc.Y, colWidths[colIdx], rowHeight);
                        rcCell.Inflate(-style.CellPadding, 0);

                        Color textColor = Color.Black;
                        if (item.CellColorOverrides.TryGetValue(colIdx, out Color overrideColor))
                        {
                            textColor = overrideColor;
                        }

                        _c1pdf.DrawString(cellText, fontItem, textColor, rcCell, GetAlignment(colIdx));
                    }

                    passHeight += rowHeight;
                    rc.Y += rowHeight;
                    dataRowIndex++;

                    // 列分隔線 (group 內的列之間, 最後一列不畫)
                    if (rIdx < group.Count - 1)
                    {
                        DrawRowSeparator(rc.Y);
                        rc.Y += separatorGap;
                        passHeight += separatorGap;
                    }
                }

                // Group 結束: 繪製合併欄
                if (passHeight > 0)
                {
                    DrawMergeColumns(group[0], groupTopY, passHeight);
                }

                // Groups 之間畫全寬分隔線 (最後一組不畫)
                if (gIdx < groups.Count - 1)
                {
                    _c1pdf.DrawLine(new GcPen(style.BorderColor, style.BorderWidth), rcPage.Left, rc.Y, rcPage.Right, rc.Y);
                    rc.Y += separatorGap;
                }
            }

            // === 表格結束: 繪製外框 ===
            _c1pdf.DrawLine(new GcPen(style.BorderColor, style.BorderWidth), rcPage.Left, rc.Y, rcPage.Right, rc.Y); // Bottom
            DrawOuterBorders(tableTopY, rc.Y);

            rc.Height = 0;
            return rc;
        }

        /// <summary>
        /// 依據合併欄的值將資料自動分組。
        /// 連續列中合併欄值相同的列會被歸入同一組。
        /// </summary>
        private List<List<PcncDetailItem>> BuildMergeGroups(List<PcncDetailItem> items, int[] mergeIndices)
        {
            List<List<PcncDetailItem>> groups = new List<List<PcncDetailItem>>();
            if (items.Count == 0) return groups;

            List<PcncDetailItem> currentGroup = new List<PcncDetailItem> { items[0] };
            string currentKey = GetMergeKey(items[0], mergeIndices);

            for (int i = 1; i < items.Count; i++)
            {
                string key = GetMergeKey(items[i], mergeIndices);
                if (key == currentKey)
                {
                    currentGroup.Add(items[i]);
                }
                else
                {
                    groups.Add(currentGroup);
                    currentGroup = new List<PcncDetailItem> { items[i] };
                    currentKey = key;
                }
            }
            groups.Add(currentGroup);

            return groups;
        }

        /// <summary>
        /// 取得合併欄的組合 Key (用於判斷分組)。
        /// </summary>
        private string GetMergeKey(PcncDetailItem item, int[] mergeIndices)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            for (int i = 0; i < mergeIndices.Length; i++)
            {
                if (i > 0) sb.Append('|');
                sb.Append(item.GetValueByIndex(mergeIndices[i]));
            }
            return sb.ToString();
        }

        /// <summary>
        /// 檢查 PDF 是否已初始化（至少有一頁，且尚未被匯出封裝）。
        /// </summary>
        public bool IsPdfInitialized => Pdf?.Pages.Count > 0 && !_isExported;

        /// <summary>
        /// 釋放資源。
        /// </summary>
        public virtual void Dispose()
        {
            _gcLogo?.Dispose();
            _logoMs?.Dispose();
            GC.SuppressFinalize(this);
        }
        /// <summary>
        /// 繪製 Flag Note 至指定 Y 位置。
        /// 包含 supplemental text（若有）以及逐行繪製 notes 清單。
        /// </summary>
        /// <param name="notes">Note 列表</param>
        /// <param name="supplementalText">補充文字</param>
        /// <param name="aboveFooter">是否添加到 Footer 上方</param>
        public virtual void DrawFlagNote(List<string> notes, string supplementalText, bool aboveFooter)
        {
            if (notes == null || notes.Count == 0)
            {
                Logger.Debug("DrawFlagNote: no notes, skip.");
                return;
            }

            float lineHeight = 14f;
            float startX = PageRect.Left;
            float drawWidth = PageRect.Width;
            float currentDrawY = CurrentRect.Y;

            C1Font flagFont = new C1Font("Arial", 8, C1.Util.FontStyle.Regular);
            C1StringFormat sfLeft = new C1StringFormat();
            sfLeft.LineAlignment = C1.Util.VerticalAlignment.Center;
            sfLeft.Alignment = C1.Util.HorizontalAlignment.Left;

            // Draw supplemental text first if exists
            if (!string.IsNullOrEmpty(supplementalText))
            {
                RectangleF rcSupp = new RectangleF(startX, currentDrawY, drawWidth, lineHeight);
                Pdf.DrawString(supplementalText, flagFont, Color.Black, rcSupp, sfLeft);
                currentDrawY += lineHeight;
            }

            // Draw each note line
            for (int i = 0; i < notes.Count; i++)
            {
                RectangleF rcNote = new RectangleF(startX, currentDrawY, drawWidth, lineHeight);
                Pdf.DrawString(notes[i], flagFont, Color.Black, rcNote, sfLeft);
                currentDrawY += lineHeight;
            }

            // Update CurrentRect
            float totalDrawn = currentDrawY - CurrentRect.Y;
            CurrentRect = new RectangleF(
                CurrentRect.X,
                currentDrawY,
                CurrentRect.Width,
                CurrentRect.Height - totalDrawn);

            Logger.Info($"DrawFlagNote done: {notes.Count} notes drawn.");
        }

        /// <summary>
        /// 補充文字的高度
        /// </summary>
        public float SupplementalTextHeight { get; set; } = 0f;

        /// <summary>
        /// 計算補充文字所需的高度
        /// </summary>
        /// <param name="supplementalText">補充說明文字</param>
        /// <returns>高度</returns>
        public virtual float CalculateSupplementalTextHeight(string supplementalText)
        {
            if (string.IsNullOrEmpty(supplementalText))
                return 0f;

            // Arial 8pt, line height 14pt
            return 14f;
        }

        /// <summary>
        /// 計算 Flag Note 所需的總高度（不包含邊界留白，僅文字與間距）
        /// </summary>
        /// <param name="notes">Flag Note 項目清單</param>
        /// <param name="supplementalText">附加的補充文字 (可為 null)</param>
        /// <returns>總高度 (pt)</returns>
        public virtual float CalculateTotalFlagNoteHeight(List<string> notes, string? supplementalText)
        {
            float defaultLineHeight = 14f;
            float totalHeight = 0f;

            if (notes != null)
            {
                totalHeight += notes.Count * defaultLineHeight;
            }

            if (!string.IsNullOrEmpty(supplementalText))
            {
                totalHeight += CalculateSupplementalTextHeight(supplementalText);
            }

            return totalHeight;
        }

        /// <summary>
        /// 繪製 Summary Result Table( 6欄 )。
        /// 支援換頁邏輯、奇偶列底色交替、多行文字自動換行。
        /// 當 drawFlagNote 為 true 時，換頁前繪製 Flag Note，且 Footer 保留高度含 Flag Note Height + 2pt。
        /// 換頁時會根據 pageBottomMargin 預留距離 Footer 頂線的緩衝空間。
        /// </summary>
        /// <param name="data">Summary 資料 (每列 6 欄: Well ID, Specimen ID, Status, Nucleotide Change, Mutation, Flag)</param>
        /// <param name="flagNoteData">Flag Note 項目清單 (可為 null)</param>
        /// <param name="supplementalText">補充文字 (可為 null)</param>
        /// <param name="drawFlagNote">是否在換頁時繪製 Flag Note 並預留高度</param>
        /// <param name="pageBottomMargin">換頁距離底部的緩衝區大小 (預設 0f)</param>
        public virtual void DrawSummaryResult6ColumnTable(
            List<List<string>> data,
            List<string>? flagNoteData,
            string? supplementalText,
            bool drawFlagNote,
            float pageBottomMargin = 0f)
        {
            if (data == null || data.Count == 0)
            {
                Logger.Info("DrawSummaryResultTable: no data, skip.");
                return;
            }

            Logger.Info($"DrawSummaryResultTable: {data.Count} rows, drawFlagNote={drawFlagNote}");

            // === Table style parameters ===
            string[] headers = { "Well ID", "Specimen ID", "Status", "Nucleotide Change", "Mutation", "Flag" };
            int numOfCol = 6;
            float fontSize = 8f;
            float cellPadding = 3f;
            float borderWidth = 0.5f;
            float flagNoteGap = FlagNoteSpacing;

            // Column width ratios (sum = 1.0)
            float[] colRatios = { 0.08f, 0.10f, 0.08f, 0.30f, 0.34f, 0.10f };

            // === Calculate column widths and X positions ===
            float tableWidth = PageRect.Width;
            float[] colWidths = new float[numOfCol];
            for (int i = 0; i < numOfCol; i++)
            {
                colWidths[i] = tableWidth * colRatios[i];
            }

            float[] colX = new float[numOfCol];
            colX[0] = PageRect.Left;
            for (int i = 1; i < numOfCol; i++)
            {
                colX[i] = colX[i - 1] + colWidths[i - 1];
            }

            // === Fonts and formats ===
            C1Font fontHeader = new C1Font("Arial", fontSize, C1.Util.FontStyle.Bold);
            C1Font fontItem = new C1Font("Arial", fontSize, C1.Util.FontStyle.Regular);

            C1StringFormat sfLeft = new C1StringFormat();
            sfLeft.LineAlignment = C1.Util.VerticalAlignment.Center;
            sfLeft.Alignment = C1.Util.HorizontalAlignment.Left;

            C1StringFormat sfCenter = new C1StringFormat();
            sfCenter.LineAlignment = C1.Util.VerticalAlignment.Center;
            sfCenter.Alignment = C1.Util.HorizontalAlignment.Center;

            // Column alignments (0=Left, 2=Center)
            int[] colAligns = { 2, 2, 2, 0, 0, 2 };
            C1StringFormat GetAlignment(int colIdx)
            {
                return colAligns[colIdx] == 2 ? sfCenter : sfLeft;
            }

            // === Calculate Flag Note reserved height ===
            float flagNoteReservedHeight = 0f;
            if (drawFlagNote && flagNoteData != null && flagNoteData.Count > 0)
            {
                flagNoteReservedHeight = CalculateTotalFlagNoteHeight(flagNoteData, supplementalText) + flagNoteGap;
            }

            // === Drawable bottom boundary ===
            float drawableBottom = PageRect.Bottom - flagNoteReservedHeight - pageBottomMargin;

            // === Alternating row colors ===
            Color evenRowColor = Color.FromArgb(242, 242, 242);
            Color oddRowColor = Color.White;

            // === NormalizeText: handle embedded newlines ===
            string NormalizeText(string text)
            {
                if (string.IsNullOrEmpty(text)) return text;
                string s = text.Replace("\\r\\n", "\n").Replace("\\n", "\n").Replace("\\r", "\n");
                return s;
            }

            // === Calculate header height ===
            float headerHeight = 0;
            for (int i = 0; i < numOfCol; i++)
            {
                float h = Pdf.MeasureString(headers[i], fontHeader, colWidths[i] - cellPadding * 2).Height + 6;
                if (h > headerHeight) headerHeight = h;
            }

            RectangleF rc = CurrentRect;

            // === Helper: Draw table header ===
            void DrawTableHeader()
            {
                Pdf.DrawLine(new GcPen(Color.Black, borderWidth), PageRect.Left, rc.Y, PageRect.Left + tableWidth, rc.Y);

                RectangleF rcHeaderBg = new RectangleF(PageRect.Left, rc.Y, tableWidth, headerHeight);
                Pdf.FillRectangle(Color.FromArgb(200, 200, 200), rcHeaderBg);

                for (int i = 0; i < numOfCol; i++)
                {
                    RectangleF rcCell = new RectangleF(colX[i], rc.Y, colWidths[i], headerHeight);
                    rcCell.Inflate(-cellPadding, 0);
                    Pdf.DrawString(headers[i], fontHeader, Color.Black, rcCell, GetAlignment(i));
                }

                rc = new RectangleF(rc.X, rc.Y + headerHeight, rc.Width, rc.Height - headerHeight);
                Pdf.DrawLine(new GcPen(Color.Black, borderWidth), PageRect.Left, rc.Y, PageRect.Left + tableWidth, rc.Y);
            }

            // === Helper: Draw Flag Note on current page ===
            void DrawFlagNoteOnPage()
            {
                if (!drawFlagNote || flagNoteData == null || flagNoteData.Count == 0) return;

                float flagStartY = drawableBottom + flagNoteGap;
                CurrentRect = new RectangleF(
                    CurrentRect.X,
                    flagStartY,
                    CurrentRect.Width,
                    PageRect.Bottom - flagStartY);
                DrawFlagNote(flagNoteData, supplementalText ?? "", true);
            }

            // === Look-ahead for first data row (Prevent Orphan Header) ===
            float firstRowHeightEstimate = 14f;
            if (data.Count > 0)
            {
                var firstRow = data[0];
                for (int c = 0; c < numOfCol; c++)
                {
                    string cellText = c < firstRow.Count ? firstRow[c] : "";
                    float cellH = Pdf.MeasureString(NormalizeText(cellText), fontItem, colWidths[c] - cellPadding * 2).Height + 4;
                    if (cellH > firstRowHeightEstimate) firstRowHeightEstimate = cellH;
                }
            }

            if (rc.Y > PageRect.Top + 10f && rc.Y + headerHeight + firstRowHeightEstimate > drawableBottom)
            {
                // Force page break before drawing header if space is not enough for first row
                CurrentRect = rc;
                DrawFlagNoteOnPage();
                AddNewPage();
                rc = CurrentRect;
                drawableBottom = PageRect.Bottom - flagNoteReservedHeight;
            }

            // === Draw Header ===
            DrawTableHeader();

            // === Draw data rows ===
            int globalRowIndex = 0;
            for (int rowIdx = 0; rowIdx < data.Count; rowIdx++)
            {
                var row = data[rowIdx];

                // Calculate row height (considering multiline text)
                float rowHeight = 0;
                string[] normalizedCells = new string[numOfCol];
                for (int c = 0; c < numOfCol; c++)
                {
                    string cellText = c < row.Count ? row[c] : "";
                    normalizedCells[c] = NormalizeText(cellText);

                    float cellH = Pdf.MeasureString(
                        normalizedCells[c],
                        fontItem,
                        colWidths[c] - cellPadding * 2).Height + 4;
                    if (cellH > rowHeight) rowHeight = cellH;
                }

                if (rowHeight < 14f) rowHeight = 14f;

                // === Page break check ===
                if (rc.Y + rowHeight > drawableBottom)
                {
                    // Draw bottom border before page break
                    Pdf.DrawLine(new GcPen(Color.Black, borderWidth), PageRect.Left, rc.Y, PageRect.Left + tableWidth, rc.Y);

                    // Draw Flag Note before page break
                    CurrentRect = rc;
                    DrawFlagNoteOnPage();

                    // New page
                    AddNewPage();
                    rc = CurrentRect;

                    // Recalculate drawable bottom for new page
                    drawableBottom = PageRect.Bottom - flagNoteReservedHeight;

                    // Redraw header on new page
                    DrawTableHeader();
                }

                // Draw alternating row background
                Color rowBgColor = (globalRowIndex % 2 == 0) ? evenRowColor : oddRowColor;
                RectangleF rcRowBg = new RectangleF(PageRect.Left, rc.Y, tableWidth, rowHeight);
                Pdf.FillRectangle(rowBgColor, rcRowBg);

                // Draw cell text
                for (int c = 0; c < numOfCol; c++)
                {
                    RectangleF rcCell = new RectangleF(colX[c], rc.Y, colWidths[c], rowHeight);
                    rcCell.Inflate(-cellPadding, 0);
                    Pdf.DrawString(normalizedCells[c], fontItem, Color.Black, rcCell, GetAlignment(c));
                }

                // Draw row separator
                Pdf.DrawLine(new GcPen(Color.LightGray, 0.2f), PageRect.Left, rc.Y + rowHeight, PageRect.Left + tableWidth, rc.Y + rowHeight);

                rc = new RectangleF(rc.X, rc.Y + rowHeight, rc.Width, rc.Height - rowHeight);
                globalRowIndex++;
            }

            // === Draw table bottom border ===
            Pdf.DrawLine(new GcPen(Color.Black, borderWidth), PageRect.Left, rc.Y, PageRect.Left + tableWidth, rc.Y);

            // === Last page: draw Flag Note below table if space allows ===
            if (drawFlagNote && flagNoteData != null && flagNoteData.Count > 0)
            {
                float spaceRemaining = drawableBottom - rc.Y;
                if (spaceRemaining > 0)
                {
                    CurrentRect = new RectangleF(
                        rc.X,
                        rc.Y + flagNoteGap,
                        rc.Width,
                        rc.Height - flagNoteGap);
                    DrawFlagNote(flagNoteData, supplementalText ?? "", true);
                }
            }
            else
            {
                CurrentRect = rc;
            }

            Logger.Info($"DrawSummaryResultTable done: {data.Count} rows.");
        }

        /// <summary>
        /// 繪製 Well Info Table。
        /// 每列上方繪製灰色分隔線。
        /// Is2Column=false: 單行 "Key: Value" 左對齊。
        /// Is2Column=true: Key 左對齊、Value 右對齊（雙欄佈局）。
        /// 支援自動換頁。
        /// </summary>
        /// <param name="items">Well Info 資料項目</param>
        public virtual void DrawWellInfoTable(List<WellInfoItem> items)
        {
            if (items == null || items.Count == 0)
            {
                Logger.Info("DrawWellInfoTable: no data, skip.");
                return;
            }

            Logger.Info($"DrawWellInfoTable: {items.Count} items.");

            // Style parameters
            float lineWidth = 0.2f;
            float rowPadding = 3f;
            Color lineColor = Color.Gray;
            C1Font textFont = BodyFont; // Arial 9pt

            C1StringFormat sfLeft = new C1StringFormat();
            sfLeft.LineAlignment = C1.Util.VerticalAlignment.Center;
            sfLeft.Alignment = C1.Util.HorizontalAlignment.Left;

            C1StringFormat sfRight = new C1StringFormat();
            sfRight.LineAlignment = C1.Util.VerticalAlignment.Center;
            sfRight.Alignment = C1.Util.HorizontalAlignment.Right;

            float drawWidth = PageRect.Width;
            float drawLeft = PageRect.Left;
            float drawRight = PageRect.Left + drawWidth;

            for (int i = 0; i < items.Count; i++)
            {
                WellInfoItem item = items[i];

                // Determine display text for measurement
                string displayText = item.Is2Column
                    ? item.Key
                    : $"{item.Key}: {item.Value}";

                // Measure row height
                float textHeight = Pdf.MeasureString(displayText, textFont, drawWidth).Height;
                float rowHeight = textHeight + rowPadding * 2;
                if (rowHeight < 16f) rowHeight = 16f;

                // Total height = line + row
                float totalNeeded = lineWidth + rowHeight;

                // Page break check
                if (CurrentRect.Y + totalNeeded > PageRect.Bottom)
                {
                    AddNewPage();
                }

                // Draw separator line above row
                Pdf.DrawLine(
                    new GcPen(lineColor, lineWidth),
                    drawLeft, CurrentRect.Y,
                    drawRight, CurrentRect.Y);

                // Draw row content
                float textY = CurrentRect.Y;
                RectangleF rcRow = new RectangleF(drawLeft, textY, drawWidth, rowHeight);

                if (item.Is2Column)
                {
                    // Two-column: Key 左半靠左, Value 從中間位置靠左對齊
                    RectangleF rcKey = new RectangleF(drawLeft + rowPadding, textY, drawWidth / 2 - rowPadding, rowHeight);
                    RectangleF rcVal = new RectangleF(drawLeft + drawWidth / 2, textY, drawWidth / 2 - rowPadding, rowHeight);
                    Pdf.DrawString(item.Key, textFont, Color.Black, rcKey, sfLeft);
                    Pdf.DrawString(item.Value, textFont, Color.Black, rcVal, sfLeft);
                }
                else
                {
                    // Single-line: "Key: Value" left-aligned
                    RectangleF rcText = new RectangleF(drawLeft + rowPadding, textY, drawWidth - rowPadding * 2, rowHeight);
                    Pdf.DrawString($"{item.Key}: {item.Value}", textFont, Color.Black, rcText, sfLeft);
                }

                // Advance Y
                float totalDrawn = rowHeight;
                CurrentRect = new RectangleF(
                    CurrentRect.X,
                    CurrentRect.Y + totalDrawn,
                    CurrentRect.Width,
                    CurrentRect.Height - totalDrawn);
            }

            // Draw final bottom line
            Pdf.DrawLine(
                new GcPen(lineColor, lineWidth),
                drawLeft, CurrentRect.Y,
                drawRight, CurrentRect.Y);

            Logger.Info($"DrawWellInfoTable done: {items.Count} items.");
        }

        /// <summary>
        /// 繪製 Sample Control Table (對應舊版 show_SampleControl_Table_V5)。
        /// 5 欄雙區域表格: 左半 3 欄 (Controls, Result, Cutoff) + 右半 2 欄 (Controls, Result)。
        /// 左/中/右三條垂直邊框線。Header 粗體。Col 1-2 右對齊。"Fail" 顯示紅色。
        /// </summary>
        /// <param name="data">表格資料 (第 0 筆為 Header，後續為 Body，每筆 5 欄)</param>
        public virtual void DrawSampleControlTable(List<List<string>> data)
        {
            if (data == null || data.Count < 2)
            {
                Logger.Info("DrawSampleControlTable: no data or header only, skip.");
                return;
            }

            Logger.Info($"DrawSampleControlTable: {data.Count - 1} body rows.");

            int numOfCol = 5;
            float fontSize = 7f;
            float cellPadding = 8f;
            float lineWidth = 0.2f;
            Color lineColor = Color.Gray;

            C1Font headerFont = new C1Font("Arial", fontSize, C1.Util.FontStyle.Bold);
            C1Font bodyFont = new C1Font("Arial", fontSize, C1.Util.FontStyle.Regular);

            C1StringFormat sfLeft = new C1StringFormat();
            sfLeft.LineAlignment = C1.Util.VerticalAlignment.Center;
            sfLeft.Alignment = C1.Util.HorizontalAlignment.Left;

            C1StringFormat sfRight = new C1StringFormat();
            sfRight.LineAlignment = C1.Util.VerticalAlignment.Center;
            sfRight.Alignment = C1.Util.HorizontalAlignment.Right;

            float pageW = PageRect.Width;
            float pageLeft = PageRect.Left;
            float pageMid = pageLeft + pageW / 2;
            float pageRight = pageLeft + pageW;

            // Column widths (matching show_SampleControl_Table_V5):
            // Left half (pageW/2): col0 = W/4, col1 = W/8, col2 = W/8
            // Right half (pageW/2): col3 = W/4, col4 = W/4
            float[] colWidths = new float[]
            {
                pageW / 4,      // col 0
                pageW / 8,      // col 1
                pageW / 8,      // col 2
                pageW / 4,      // col 3
                pageW / 4       // col 4
            };

            float[] colX = new float[]
            {
                0,                              // col 0: start at 0
                pageW / 4,                      // col 1: start at W/4
                pageW / 4 + pageW / 8,          // col 2: start at W/4 + W/8 = 3W/8
                pageW / 2,                      // col 3: start at W/2
                pageW / 2 + pageW / 4           // col 4: start at 3W/4
            };

            float tableTop = CurrentRect.Y;
            float tableHeight = 0;

            // === Draw Header Row ===
            List<string> headerRow = data[0];
            float headerRowHeight = 0;

            for (int c = 0; c < numOfCol && c < headerRow.Count; c++)
            {
                SizeF sz = Pdf.MeasureString(headerRow[c], headerFont, colWidths[c] - cellPadding * 2);
                float h = sz.Height + 6;
                if (h > headerRowHeight) headerRowHeight = h;
            }

            // Draw header top line
            Pdf.DrawLine(new GcPen(lineColor, lineWidth), pageLeft, CurrentRect.Y, pageRight, CurrentRect.Y);

            for (int c = 0; c < numOfCol && c < headerRow.Count; c++)
            {
                RectangleF rcCell = new RectangleF(pageLeft + colX[c], CurrentRect.Y, colWidths[c], headerRowHeight);
                rcCell.Inflate(-cellPadding, 0);
                Pdf.DrawString(headerRow[c], headerFont, Color.Black, rcCell, sfLeft);
            }

            CurrentRect = new RectangleF(CurrentRect.X, CurrentRect.Y + headerRowHeight, CurrentRect.Width, CurrentRect.Height - headerRowHeight);
            tableHeight += headerRowHeight;

            // Draw header bottom line
            Pdf.DrawLine(new GcPen(lineColor, lineWidth), pageLeft, CurrentRect.Y, pageRight, CurrentRect.Y);

            // === Draw Body Rows ===
            for (int r = 1; r < data.Count; r++)
            {
                List<string> row = data[r];
                float rowHeight = 0;

                // Measure row height
                for (int c = 0; c < numOfCol && c < row.Count; c++)
                {
                    SizeF sz = Pdf.MeasureString(row[c], bodyFont, colWidths[c] - cellPadding * 2);
                    float h = sz.Height + 6;
                    if (h > rowHeight) rowHeight = h;
                }

                // Draw cells
                for (int c = 0; c < numOfCol && c < row.Count; c++)
                {
                    RectangleF rcCell = new RectangleF(pageLeft + colX[c], CurrentRect.Y, colWidths[c], rowHeight);
                    rcCell.Inflate(-cellPadding, 0);

                    // Determine alignment: col 1-2 (Result/Cutoff in left half) → right-aligned
                    C1StringFormat sf = (c == 1 || c == 2) ? sfRight : sfLeft;

                    // Determine color: col 4 "Fail" → red
                    Color textColor = Color.Black;
                    if (c == 4 && string.Equals(row[c]?.Trim(), "Fail", StringComparison.OrdinalIgnoreCase))
                    {
                        textColor = Color.Red;
                    }

                    Pdf.DrawString(row[c] ?? "", bodyFont, textColor, rcCell, sf);
                }

                CurrentRect = new RectangleF(CurrentRect.X, CurrentRect.Y + rowHeight, CurrentRect.Width, CurrentRect.Height - rowHeight);
                tableHeight += rowHeight;
            }

            // === Draw table bottom line ===
            float tableBottom = tableTop + tableHeight;
            Pdf.DrawLine(new GcPen(lineColor, lineWidth), pageLeft, tableBottom, pageRight, tableBottom);

            // === Draw vertical border lines (left, mid, right) ===
            Pdf.DrawLine(new GcPen(lineColor, lineWidth), pageLeft, tableTop, pageLeft, tableBottom);
            Pdf.DrawLine(new GcPen(lineColor, lineWidth), pageMid, tableTop, pageMid, tableBottom);
            Pdf.DrawLine(new GcPen(lineColor, lineWidth), pageRight, tableTop, pageRight, tableBottom);

            Logger.Info($"DrawSampleControlTable done: {data.Count - 1} body rows.");
        }

        /// <summary>
        /// 繪製 Individual Result Table (對應舊版 show_sample_Table_IndvPDF_V5)。
        /// 5 欄自訂寬度表格: Col 0-2 Arial 12pt 左對齊, Col 3-4 Arial 10pt 右對齊。
        /// 偶數列淡藍底色。"Detected" 紅字。換頁時重繪 Header。
        /// 換頁時會根據 pageBottomMargin 預留距離 Footer 頂線的緩衝空間。
        /// </summary>
        /// <param name="data">表格資料 (第 0 筆為 Header，後續為 Body，每筆 5 欄)</param>
        /// <param name="pageBottomMargin">換頁距離底部的緩衝區大小 (預設繼承自 PdfGlobalConfig.DefaultPageBottomMargin)</param>
        public virtual void DrawIndividualResultTable5Col(List<List<string>> data, float? pageBottomMargin = null)
        {
            if (data == null || data.Count < 2)
            {
                Logger.Info("DrawIndividualResultTable5Col: no data or header only, skip.");
                return;
            }

            Logger.Info($"DrawIndividualResultTable5Col: {data.Count - 1} body rows.");

            float bottomMargin = pageBottomMargin ?? Infrastructure.PdfGlobalConfig.DefaultPageBottomMargin;
            float drawableBottom = PageRect.Bottom - bottomMargin;

            int numOfCol = 5;
            float cellPadding = 8f;
            float lineWidth = 0.2f;
            Color lineColor = Color.Gray;
            Color altRowColor = Color.FromArgb(240, 240, 255);

            C1Font headerFont = new C1Font("Arial", 10, C1.Util.FontStyle.Bold);
            C1Font bodyFontLarge = new C1Font("Arial", 12, C1.Util.FontStyle.Regular);
            C1Font bodyFontSmall = new C1Font("Arial", 10, C1.Util.FontStyle.Regular);

            // 將 CSV 中的字面 \r\n 轉換為實際換行符
            string NormalizeLineBreaks(string text)
            {
                if (string.IsNullOrEmpty(text)) return text;
                return text.Replace("\\r\\n", "\n").Replace("\\n", "\n").Replace("\\r", "\n");
            }

            // 強制以字元級別過長斷行，避免 PDF 繪圖將標點當作換行點導致提早斷行
            string WrapTextCharByChar(string text, C1Font font, float maxWidth)
            {
                if (string.IsNullOrEmpty(text)) return text;
                string[] paragraphs = text.Split(new[] { '\n' }, StringSplitOptions.None);
                List<string> finalLines = new List<string>();

                foreach (string p in paragraphs)
                {
                    if (string.IsNullOrEmpty(p))
                    {
                        finalLines.Add("");
                        continue;
                    }

                    string currentLine = "";
                    for (int i = 0; i < p.Length; i++)
                    {
                        string testLine = currentLine + p[i];
                        SizeF sz = Pdf.MeasureString(testLine, font);
                        if (sz.Width > maxWidth && currentLine.Length > 0)
                        {
                            finalLines.Add(currentLine);
                            currentLine = p[i].ToString();
                        }
                        else
                        {
                            currentLine = testLine;
                        }
                    }
                    if (currentLine.Length > 0)
                    {
                        finalLines.Add(currentLine);
                    }
                }
                return string.Join("\n", finalLines);
            }

            C1StringFormat sfLeft = new C1StringFormat();
            sfLeft.LineAlignment = C1.Util.VerticalAlignment.Center;
            sfLeft.Alignment = C1.Util.HorizontalAlignment.Left;

            C1StringFormat sfRight = new C1StringFormat();
            sfRight.LineAlignment = C1.Util.VerticalAlignment.Center;
            sfRight.Alignment = C1.Util.HorizontalAlignment.Right;

            float pageW = PageRect.Width;
            float pageLeft = PageRect.Left;
            float pageRight = pageLeft + pageW;

            // Column widths (matching show_sample_Table_IndvPDF_V5)
            float[] colWidths = new float[]
            {
                pageW / 3 - 11,   // col 0: Nucleotide Change
                pageW / 3 + 10,   // col 1: Mutation
                pageW / 9 + 29,   // col 2: Result
                pageW / 9 - 17,   // col 3: MFI
                pageW / 9 - 11    // col 4: Cutoff
            };

            // Pre-compute column X positions
            float[] colX = new float[numOfCol];
            colX[0] = 0;
            for (int i = 1; i < numOfCol; i++)
            {
                colX[i] = colX[i - 1] + colWidths[i - 1];
            }

            float tableTop = CurrentRect.Y;
            float tableHeight = 0;
            List<string> headerRow = data[0];

            // === Helper: Draw Header ===
            void DrawHeader()
            {
                float rowHeight = 0;
                for (int c = 0; c < numOfCol && c < headerRow.Count; c++)
                {
                    float measWidth = colWidths[c] - (c == 1 ? 2 : 4);
                    string headerText = WrapTextCharByChar(NormalizeLineBreaks(headerRow[c] ?? ""), headerFont, measWidth);
                    SizeF sz = Pdf.MeasureString(headerText, headerFont, measWidth);
                    float h = sz.Height + 6;
                    if (h > rowHeight) rowHeight = h;
                }

                // Header top line
                Pdf.DrawLine(new GcPen(lineColor, lineWidth), pageLeft, CurrentRect.Y, pageRight, CurrentRect.Y);

                for (int c = 0; c < numOfCol && c < headerRow.Count; c++)
                {
                    float measWidth = colWidths[c] - (c == 1 ? 2 : 4);
                    RectangleF rcCell = new RectangleF(pageLeft + colX[c], CurrentRect.Y, colWidths[c], rowHeight);
                    rcCell.Inflate(-cellPadding, 0);
                    string headerText = WrapTextCharByChar(NormalizeLineBreaks(headerRow[c] ?? ""), headerFont, measWidth);
                    Pdf.DrawString(headerText, headerFont, Color.Black, rcCell, sfLeft);
                }

                CurrentRect = new RectangleF(CurrentRect.X, CurrentRect.Y + rowHeight, CurrentRect.Width, CurrentRect.Height - rowHeight);
                tableHeight += rowHeight;

                // Header bottom line
                Pdf.DrawLine(new GcPen(lineColor, lineWidth), pageLeft, CurrentRect.Y, pageRight, CurrentRect.Y);
            }

            // === Look-ahead for first data row (Prevent Orphan Header) ===
            float headerHeightEstimate = 0f;
            for (int c = 0; c < numOfCol && c < headerRow.Count; c++)
            {
                float measWidth = colWidths[c] - (c == 1 ? 2 : 4);
                string headerText = WrapTextCharByChar(NormalizeLineBreaks(headerRow[c] ?? ""), headerFont, measWidth);
                float h = Pdf.MeasureString(headerText, headerFont, measWidth).Height + 6;
                if (h > headerHeightEstimate) headerHeightEstimate = h;
            }

            float firstRowHeightEstimate = 14f;
            if (data.Count > 1)
            {
                var firstRow = data[1];
                for (int c = 0; c < numOfCol && c < firstRow.Count; c++)
                {
                    C1Font font = (c < 3) ? bodyFontLarge : bodyFontSmall;
                    float measWidth = colWidths[c] - (c <= 1 ? 2 : 4);
                    string cellText = WrapTextCharByChar(NormalizeLineBreaks(firstRow[c] ?? ""), font, measWidth);
                    float h = Pdf.MeasureString(cellText, font, measWidth).Height + 6;
                    if (h > firstRowHeightEstimate) firstRowHeightEstimate = h;
                }
            }

            bool initialPageBreak = CurrentRect.Y > PageRect.Top + 10f && CurrentRect.Y + headerHeightEstimate + firstRowHeightEstimate > drawableBottom;
            Logger.Info($"[INDV Table] Initial Page Break Check: CurrentY={CurrentRect.Y:F2}, PageTop={PageRect.Top:F2}, HeaderEst={headerHeightEstimate:F2}, FirstRowEst={firstRowHeightEstimate:F2}, DrawableBottom={drawableBottom:F2}. Triggered: {initialPageBreak}");
            if (initialPageBreak)
            {
                AddNewPage();
            }

            // === Draw initial Header ===
            DrawHeader();

            // === Draw Body Rows ===
            for (int r = 1; r < data.Count; r++)
            {
                List<string> row = data[r];
                float rowHeight = 0;

                // Pre-compute row height
                for (int c = 0; c < numOfCol && c < row.Count; c++)
                {
                    C1Font font = (c < 3) ? bodyFontLarge : bodyFontSmall;
                    float measWidth = colWidths[c] - (c <= 1 ? 2 : 4);
                    string cellText = WrapTextCharByChar(NormalizeLineBreaks(row[c] ?? ""), font, measWidth);
                    SizeF sz = Pdf.MeasureString(cellText, font, measWidth);
                    float h = sz.Height + 6;
                    if (h > rowHeight) rowHeight = h;
                }

                // === Page break check (BEFORE drawing — consistent with other tables) ===
                bool rowPageBreak = CurrentRect.Y + rowHeight > drawableBottom;
                if (rowPageBreak)
                {
                    float distanceToBottom = PageRect.Bottom - CurrentRect.Y;
                    Logger.Info($"[INDV Table] Row Page Break (Row {r}): CurrentY={CurrentRect.Y:F2}, RowHeight={rowHeight:F2}, DrawableBottom={drawableBottom:F2}, DistanceToBottom={distanceToBottom:F2}, Buffer={bottomMargin:F2}. (因為 Y+Height > DrawableBottom 觸發換頁)");
                    
                    // Draw bottom line + vertical borders for current page segment
                    Pdf.DrawLine(new GcPen(lineColor, lineWidth), pageLeft, CurrentRect.Y, pageRight, CurrentRect.Y);
                    float segBottom = tableTop + tableHeight;
                    Pdf.DrawLine(new GcPen(lineColor, lineWidth), pageLeft, tableTop, pageLeft, segBottom);
                    Pdf.DrawLine(new GcPen(lineColor, lineWidth), pageRight, tableTop, pageRight, segBottom);

                    // New page
                    AddNewPage();
                    CurrentRect = new RectangleF(PageRect.X, PageRect.Top + 20, PageRect.Width, PageRect.Height - 20);
                    tableTop = CurrentRect.Y;
                    tableHeight = 0;

                    // Re-draw header on new page
                    DrawHeader();
                }

                // Even-row alternate background
                if (r % 2 == 0)
                {
                    Pdf.FillRectangle(altRowColor, pageLeft, CurrentRect.Y, pageW, rowHeight);
                }

                // Draw cells
                for (int c = 0; c < numOfCol && c < row.Count; c++)
                {
                    C1Font font = (c < 3) ? bodyFontLarge : bodyFontSmall;
                    C1StringFormat sf = (c < 3) ? sfLeft : sfRight;
                    float inflate = (c <= 1) ? 1 : 2;

                    RectangleF rcCell = new RectangleF(pageLeft + colX[c], CurrentRect.Y, colWidths[c], rowHeight);
                    rcCell.Inflate(-inflate, 0);

                    float measWidth = colWidths[c] - (c <= 1 ? 2 : 4);
                    string cellText = WrapTextCharByChar(NormalizeLineBreaks(row[c] ?? ""), font, measWidth);
                    
                    // Red text for "Detected"
                    Color textColor = Color.Black;
                    if (string.Equals(row[c]?.Trim(), "Detected", StringComparison.OrdinalIgnoreCase))
                    {
                        textColor = Color.Red;
                    }

                    Pdf.DrawString(cellText, font, textColor, rcCell, sf);
                }

                CurrentRect = new RectangleF(CurrentRect.X, CurrentRect.Y + rowHeight, CurrentRect.Width, CurrentRect.Height - rowHeight);
                tableHeight += rowHeight;
            }

            // === Draw table bottom line + vertical borders ===
            float finalBottom = tableTop + tableHeight;
            Pdf.DrawLine(new GcPen(lineColor, lineWidth), pageLeft, finalBottom, pageRight, finalBottom);
            Pdf.DrawLine(new GcPen(lineColor, lineWidth), pageLeft, tableTop, pageLeft, finalBottom);
            Pdf.DrawLine(new GcPen(lineColor, lineWidth), pageRight, tableTop, pageRight, finalBottom);

            Logger.Info($"DrawIndividualResultTable5Col done: {data.Count - 1} body rows.");
        }
    }
}

