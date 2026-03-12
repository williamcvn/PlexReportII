# PlexReportII

PlexReportII 是一個基於 WinForms 與 ComponentOne (C1.Pdf) 的 PDF 報表產生核心函式庫。
此專案封裝了繪製精美 PDF 醫療檢驗報告（如 EGFR Plus Assay Report）的複雜邏輯，提供各類表格、文字段落、與排版功能的 API。

## 系統需求
- **目標框架**: .NET 8.0
- **第三方依賴**:
  - `System.Drawing.Common` (v10.0.2)
  - `ComponentOne C1.Pdf` (v10.0.20252.203) - **需要有效商業授權**

## 授權注意
本專案為內部私有專案，包含需付費授權的元件（[SOUP] ComponentOne C1.Pdf）。
請勿將專案設為 Public，且在不同電腦建置前，請確保已安裝 GrapeCity License Manager 並啟用合法的金鑰，否則編譯出的 PDF 會無法產生或帶有浮水印。
