# Google Apps Script 專案自動化初始化腳本 (Windows PowerShell)

Write-Host "--- 開始掃描並初始化所有 GAS 專案 ---" -ForegroundColor Cyan

# 1. 檢查是否安裝 clasp
if (-not (Get-Command "clasp" -ErrorAction SilentlyContinue)) {
    Write-Error "找不到 clasp 指令，請先執行 'npm install -g @google/clasp'"
    exit
}

# 2. 獲取所有含有 code.gs 的資料夾
$projects = Get-ChildItem -Path . -Filter "code.gs" -Recurse | ForEach-Object { $_.DirectoryName } | Select-Object -Unique

foreach ($folder in $projects) {
    $folderName = Split-Path $folder -Leaf
    $claspJsonPath = Join-Path $folder ".clasp.json"
    
    Write-Host "`n[正在處理] $folderName" -ForegroundColor Yellow
    
    Push-Location $folder
    
    # 檢查是否需要建立新專案
    # 狀況：沒有 .clasp.json 或是 scriptId 為空
    $needCreate = $false
    if (-not (Test-Path ".clasp.json")) {
        $needCreate = $true
    } else {
        $claspJson = Get-Content ".clasp.json" | ConvertFrom-Json
        if ([string]::IsNullOrWhiteSpace($claspJson.scriptId)) {
            $needCreate = $true
        }
    }
    
    if ($needCreate) {
        Write-Host "  -> 偵測到未綁定，正在建立全新的雲端試算表專案..." -ForegroundColor Gray
        # 建立試算表類型的 GAS 專案
        clasp create --type sheets --title "Tutorial_$folderName"
    } else {
        Write-Host "  -> 已有綁定 ID: $($claspJson.scriptId)" -ForegroundColor Green
    }
    
    # 推送程式碼
    Write-Host "  -> 正在推送程式碼到雲端..." -ForegroundColor Gray
    clasp push -f
    
    Pop-Location
}

Write-Host "`n✅ 所有專案處理完成！" -ForegroundColor Green
Write-Host "現在你可以去 Google Drive 查看這些專案了。"
