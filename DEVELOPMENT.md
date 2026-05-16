# 開發者與部署指南 (Windows PowerShell)

本專案使用 [clasp](https://github.com/google/clasp) 工具搭配 PowerShell 腳本進行自動化部署。為了保護隱私，`.clasp.json` (包含 scriptId) 已被列入忽略清單，不會上傳至 GitHub。

## 🛠 環境建置 (僅需執行一次)

1. **安裝 Node.js** (建議 v22 以上版本)。
2. **安裝 clasp**:
   ```powershell
   npm install -g @google/clasp
   ```
3. **登入 Google 帳號**:
   ```powershell
   clasp login
   ```
4. **啟用 Apps Script API**:
   前往 [Google Apps Script 使用者設定](https://script.google.com/home/usersettings) 並將其切換為 **ON**。

---

## 🚀 自動化部署流程

當你初次 `git clone` 或是想要部署某個 Session 時，請執行以下步驟：

### 1. 檢查目前部署狀態
執行以下指令可以查看哪些 Session 尚未綁定雲端專案：
```powershell
Get-ChildItem -Path . -Filter "code.gs" -Recurse | ForEach-Object {
    $folder = $_.DirectoryName
    $status = Test-Path "$folder\.clasp.json"
    $id = if ($status) { (Get-Content "$folder\.clasp.json" | ConvertFrom-Json).scriptId } else { "" }
    
    $color = if ($id -eq "") { "Yellow" } else { "Green" }
    Write-Host ("[{0}] {1}" -f (if ($id -eq "") { "未綁定" } else { "已綁定" }), $_.FullName) -ForegroundColor $color
}
```

### 2. 一鍵自動建立雲端專案 (PowerShell 魔法)
如果你想針對某個資料夾（例如 `Session01_GAS基本介面與巨集`）進行自動化部署（自動建立試算表 + 自動綁定），請執行：

```powershell
# 請修改下方變數為你要部署的資料夾名稱
$TARGET_FOLDER = "Session01_GAS基本介面與巨集"

# 執行自動部署
Set-Location $TARGET_FOLDER
clasp create --type sheets --title "$TARGET_FOLDER"
clasp push -f
Write-Host "✅ 部署完成！請至 Google Drive 查看新建立的試算表。" -ForegroundColor Green
Set-Location ..
```

---

## 🔒 關於安全性

- **`.clasp.json`**: 已在 `.gitignore` 中過濾，確保你的個人 `scriptId` 不會洩漏。
- **`.clasprc.json`**: 這是你的登入憑證，**絕對不可**上傳至任何公開空間。
