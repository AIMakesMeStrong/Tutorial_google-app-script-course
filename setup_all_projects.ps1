# Google Apps Script Project Automation (Final Working Version)

Write-Host "--- Starting GAS Project Initialization ---" -ForegroundColor Cyan

# 1. Check for clasp
if (-not (Get-Command "clasp" -ErrorAction SilentlyContinue)) {
    Write-Error "clasp not found. Please run: npm install -g @google/clasp"
    exit
}

# 2. Find ALL code.gs files
Write-Host "Scanning for code.gs files..." -ForegroundColor Gray
$gsFiles = Get-ChildItem -Path $PSScriptRoot -Filter "code.gs" -Recurse -File

if ($null -eq $gsFiles -or $gsFiles.Count -eq 0) {
    Write-Warning "No code.gs files found."
    exit
}

Write-Host "Found $($gsFiles.Count) potential projects." -ForegroundColor Green

foreach ($file in $gsFiles) {
    $folder = $file.DirectoryName
    $folderName = $file.Directory.Name
    
    Write-Host "`n[Processing] $folderName" -ForegroundColor Yellow
    Write-Host "  Location: $folder" -ForegroundColor DarkGray
    
    Push-Location -LiteralPath $folder
    
    try {
        # Determine if we need to create a new cloud project
        $needCreate = $false
        if (-not (Test-Path ".clasp.json")) {
            $needCreate = $true
        } else {
            $content = Get-Content ".clasp.json" -Raw -ErrorAction SilentlyContinue
            if ([string]::IsNullOrWhiteSpace($content)) {
                $needCreate = $true
            } else {
                $claspJson = $content | ConvertFrom-Json
                if ([string]::IsNullOrWhiteSpace($claspJson.scriptId)) {
                    $needCreate = $true
                }
            }
        }
        
        if ($needCreate) {
            Write-Host "  -> Creating new Google Sheet project..." -ForegroundColor Gray
            
            # FIX: clasp searches PARENT directories for .clasp.json
            # So we create the project in a clean temp folder, then move the result back
            $tempDir = Join-Path $env:TEMP ("clasp_temp_" + [guid]::NewGuid().ToString("N"))
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
            
            Push-Location -LiteralPath $tempDir
            clasp create --type sheets --title "Tutorial_$folderName"
            Pop-Location
            
            # Move the generated .clasp.json back to the target folder
            $tempClasp = Join-Path $tempDir ".clasp.json"
            if (Test-Path $tempClasp) {
                # Remove old empty .clasp.json if it exists
                if (Test-Path ".clasp.json") { Remove-Item ".clasp.json" -Force }
                Copy-Item $tempClasp ".clasp.json" -Force
                Write-Host "  -> Project created successfully!" -ForegroundColor Green
            } else {
                Write-Warning "  -> clasp create did not produce .clasp.json"
            }
            
            # Clean up temp folder
            Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            
        } else {
            Write-Host "  -> Using existing ID: $($claspJson.scriptId)" -ForegroundColor Green
        }
        
        # Enforce Asia/Taipei TimeZone
        if (Test-Path "appsscript.json") {
            $manifest = Get-Content "appsscript.json" -Raw | ConvertFrom-Json
            $manifest.timeZone = "Asia/Taipei"
            $manifest | ConvertTo-Json | Set-Content "appsscript.json"
        }
        
        # Push code to cloud
        Write-Host "  -> Pushing code..." -ForegroundColor Gray
        clasp push -f
        
    } catch {
        Write-Warning "  Error in $folderName : $($_.Exception.Message)"
    } finally {
        Pop-Location
    }
}

Write-Host "`n All done!" -ForegroundColor Green
