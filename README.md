# Auto-Adjust Word Format

自動調整 Word (.docx) 文件格式的工具，針對特定表格結構進行批次修改。

## 功能

1. **刪除欄位** — 移除表格中的 "No. of Samples" 欄位（若存在）
2. **拆分欄位** — 將 "Result / Rating" 拆成獨立的 "Result" 與 "Rating" 兩欄
3. **統一字型** — 全文改為 Tahoma, 10pt
4. **表頭清理** — 去除表頭多餘字元，統一置中對齊；資料表第一列（標題列）固定列高
5. **版面轉換** — 橫式轉直式，表格寬度設為頁面可印寬度（滿版）
6. **避免表格超寬** — 若 `tblGrid` 欄寬總和超過可印寬度，依比例縮放（保留各欄相對寬度），並同步資料表的儲存格 `tcW`

## 使用方式

### 命令列

```bash
# 建立虛擬環境
python3 -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt

# 執行（預設讀 source/ 輸出至 target/）
python format_docx.py

# 自訂輸入/輸出路徑
python format_docx.py input.docx output.docx
```

### RTF 檔（`.rtf`）

使用本機 **Microsoft Word**（透過 `pywin32` COM）將 RTF 存成 `.docx`，再執行與 `.docx` 相同的格式調整。需已安裝 Word 且 Python 環境已正確安裝並註冊 `pywin32`。

```bash
python format_docx.py source/report.rtf target/report_adjusted.docx
```

若只給 RTF 路徑、未指定輸出，預設為 `target/<檔名>_adjusted.docx`。中間會在輸出目錄產生 `<檔名>_converted.docx`。

### Windows EXE

雙擊 `FormatDocx.exe` → 選擇 `.docx` 或 `.rtf` → 轉換後的檔案會存在原檔同目錄，檔名加上 `_adjusted`。

#### 打包 EXE

**方式一：本機打包（需 Windows + Python 3.10+）**

```
雙擊 build_exe.bat
產出：dist\FormatDocx.exe
```

**方式二：GitHub Actions**

推送 tag 即自動建置：

```bash
git tag v1.0
git push --tags
```

至 Actions 頁面下載 `FormatDocx-windows` artifact。

## 專案結構

```
├── format_docx.py        # 核心轉換邏輯
├── rtf_to_docx.py        # RTF→DOCX（Word COM）
├── format_docx_gui.py    # GUI 包裝（檔案選擇對話框）
├── build_exe.bat         # Windows 一鍵打包腳本
├── requirements.txt      # Python 依賴
└── .github/workflows/    # GitHub Actions CI
```
