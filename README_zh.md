# 蝦皮訂單爬蟲

[English](README.md) | [中文](README_zh.md)

一個功能強大的智能網頁爬蟲，用於從蝦皮賣家中心提取訂單數據。此工具會自動登入、瀏覽訂單，並將詳細的訂單資訊匯出為 Excel 格式。

## 功能特色

- **智能登入系統**：基於 Cookie 的身份驗證，自動保存登入狀態
- **反偵測瀏覽器**：使用 undetected-chromedriver 避免被偵測
- **穩健的元素互動**：多種備用策略確保元素點擊成功
- **全面的數據提取**：提取訂單詳情、買家資訊、配送方式和商品規格
- **Excel 匯出**：自動格式化並匯出數據到 Excel，包含適當的樣式
- **錯誤處理**：全面的錯誤處理和重試機制
- **跨平台支援**：支援 macOS

## 系統需求

- Python 3.8 或更高版本
- Google Chrome 瀏覽器
- 蝦皮賣家帳號

## 安裝步驟

1. **複製專案**

   ```bash
   git clone <your-repository-url>
   cd crawler
   ```

2. **建立虛擬環境（建議）**

   ```bash
   python -m venv .venv
   source .venv/bin/activate  # Windows: .venv\Scripts\activate
   ```

3. **安裝依賴套件**

   ```bash
   pip install -r requirements.txt
   ```

4. **確認 Chrome 安裝**
   - 確保系統已安裝 Google Chrome
   - 爬蟲會自動下載適當版本的 ChromeDriver

## 使用方式

### 基本使用

1. **執行主爬蟲程式**

   ```bash
   python _shopee_order_crawler.py
   ```

2. **首次設定**

   - 爬蟲會開啟瀏覽器視窗
   - 如果未登入，會提示您手動登入蝦皮賣家中心
   - 成功登入後，您的登入狀態會被保存供日後使用

3. **自動數據提取**
   - 爬蟲會自動導航到訂單頁面
   - 提取所有可用的訂單數據
   - 將結果匯出到您下載資料夾中的 Excel 檔案

### 進階使用

#### 測試設定

```bash
python minimal_test.py
```

#### 執行計算工具

```bash
python Calculate.py
```

## 數據結構

爬蟲會提取每個訂單的以下資訊：

| 欄位     | 描述                          |
| -------- | ----------------------------- |
| 日期     | 訂單建立日期                  |
| 買家 ID  | 客戶用戶名                    |
| 配送方式 | 運送方式（蝦皮、7-11、OK 等） |
| 商品名稱 | 商品名稱（已過濾表情符號）    |
| 顏色     | 商品顏色規格                  |
| 尺寸     | 商品尺寸規格                  |
| 數量     | 訂購的商品數量                |
| 總價     | 訂單總金額                    |

## 設定

### 主要設定

編輯 `_shopee_order_crawler.py` 中的 `config_dict`：

```python
config_dict = {
    "advanced": {
        "verbose": True,      # 啟用除錯訊息
        "headless": False     # 無頭模式執行瀏覽器
    }
}
```

### Cookie 管理

- Cookie 會自動保存到 `shopee_cookies.pkl`
- 清除已保存的 Cookie：`python -c "from cookie_manager import CookieManager; CookieManager().clear_cookies()"`

## 專案結構

```
crawler/
├── _shopee_order_crawler.py    # 主爬蟲腳本
├── cookie_manager.py           # Cookie 管理系統
├── minimal_test.py            # 簡單測試腳本
├── Calculate.py               # 計算工具腳本
├── requirements.txt           # Python 依賴套件
├── .gitignore                # Git 忽略規則
├── webdriver/                # ChromeDriver 目錄
│   └── chromedriver         # ChromeDriver 執行檔
└── README_zh.md             # 本檔案
```

## 安全功能

### 元素點擊

爬蟲實作了多種元素點擊策略：

1. 直接點擊
2. JavaScript 點擊
3. ActionChains 點擊
4. 移除遮罩後點擊
5. 滾動到元素後點擊

### 錯誤恢復

- 自動重試機制
- 優雅處理缺失元素
- 超時管理
- 瀏覽器崩潰恢復

## 🔍 故障排除

### 常見問題

**1. ChromeDriver 版本不匹配**

```bash
# 清除 ChromeDriver 快取
rm -rf ~/.undetected_chromedriver
# 重新執行爬蟲
```

**2. 登入問題**

```bash
# 清除已保存的 Cookie
python -c "from cookie_manager import CookieManager; CookieManager().clear_cookies()"
```

**3. 找不到元素錯誤**

- 確保您已登入蝦皮賣家中心
- 檢查頁面結構是否已變更
- 嘗試在非無頭模式下執行以進行除錯

**4. 權限錯誤**

```bash
# 使 ChromeDriver 可執行（Linux/macOS）
chmod +x webdriver/chromedriver
```

### 除錯模式

在設定字典中設定 `"verbose": True` 來啟用詳細記錄。

## 📝 輸出結果

爬蟲會在您的下載資料夾中產生 Excel 檔案，命名規則如下：

- 格式：`MM-DD.xlsx`（例如：`12-25.xlsx`）
- 位置：`~/Downloads/` 目錄
- 特色：合併儲存格、格式化樣式、置中對齊

## 🤝 貢獻

1. Fork 專案
2. 建立功能分支
3. 進行您的修改
4. 徹底測試
5. 提交 Pull Request

## 📄 授權

此專案僅供教育和個人使用。請尊重蝦皮的服務條款並負責任地使用。

## ⚠️ 免責聲明

- 此工具僅供教育用途
- 請負責任地使用並遵守蝦皮的服務條款
- 作者不對任何濫用行為負責
- 請尊重速率限制和網站政策

## 🆘 支援

如果您遇到問題：

1. 檢查故障排除章節
2. 啟用除錯模式以獲得詳細記錄
3. 確保所有依賴套件都已正確安裝
4. 驗證您的蝦皮賣家帳號存取權限

---

**注意**：此爬蟲是為當前的蝦皮賣家中心介面設計的。如果蝦皮更新其網站結構，爬蟲可能需要更新以維持相容性。
