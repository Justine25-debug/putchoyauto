<# :
@echo off
setlocal

:: =====================================================================
:: COMMAND PROMPT SECTION (The Wrapper)
:: =====================================================================
:: 1. Check for Administrator privileges in CMD
net session >nul 2>&1
if %errorLevel% NEQ 0 (
    echo.
    echo [WARNING] Please right-click this file and select "Run as Administrator".
    echo.
    pause
    exit /b
)

:: 2. Check for PowerShell and install if missing
echo [INFO] Checking for PowerShell installation...

set "PS_EXE=powershell"
where powershell >nul 2>&1
if %errorLevel% EQU 0 (
    echo [INFO] Windows PowerShell detected.
    goto :LaunchScript
)

set "PS_EXE=pwsh"
where pwsh >nul 2>&1
if %errorLevel% EQU 0 (
    echo [INFO] PowerShell 7 (pwsh.exe) detected.
    goto :LaunchScript
)

echo [WARNING] No PowerShell detected on this system!
echo [INFO] Downloading PowerShell 7.6.0 from the web...

set "PS_URL=https://github.com/PowerShell/PowerShell/releases/download/v7.6.0/PowerShell-7.6.0-win-x64.msi"
set "PS_MSI=%TEMP%\PowerShell-7.6.0-win-x64.msi"

:: Download using curl (native in Windows 10/11)
curl -L -# -o "%PS_MSI%" "%PS_URL%"
if %errorLevel% NEQ 0 (
    echo [ERROR] Failed to download PowerShell MSI. Please check your internet connection.
    pause
    exit /b
)

echo [INFO] Download complete.
echo [INFO] Installing PowerShell quietly. Please wait...
start /wait msiexec.exe /i "%PS_MSI%" /quiet /qn /norestart
echo [INFO] Installation complete.

:: Clean up the downloaded installer
if exist "%PS_MSI%" del /q "%PS_MSI%"

:: Use absolute path since PATH environment variable isn't updated in the current CMD session
set "PS_EXE=C:\Program Files\PowerShell\7\pwsh.exe"
if not exist "%PS_EXE%" (
    echo [ERROR] Could not locate pwsh.exe after installation.
    pause
    exit /b
)

:LaunchScript
:: 3. Execute the PowerShell portion below while bypassing Execution Policies
echo [INFO] Launching script engine...
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Expression (Get-Content '%~f0' -Raw)"

:: 4. Keep the CMD window open after the script finishes
echo.
pause
exit /b
#>

# =====================================================================
# POWERSHELL SECTION (The Engine)
# =====================================================================

# --- Helper Function to Bypass Google Drive Virus Scan Warnings ---
function Invoke-GDriveDownload {
    param (
        [Parameter(Mandatory=$true)][string]$FileId,
        [Parameter(Mandatory=$true)][string]$OutFilePath
    )
    $BaseUri = "https://drive.google.com/uc?export=download&id=$FileId"
    $requestTimeoutSec = 45
    $Session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $TempPath = "$OutFilePath.tmp"

    function Test-ZipSignature {
        param([Parameter(Mandatory=$true)][string]$Path)
        if (-not (Test-Path $Path)) { return $false }
        $fileInfo = Get-Item -Path $Path -ErrorAction SilentlyContinue
        if (-not $fileInfo -or $fileInfo.Length -lt 2) { return $false }

        $fs = $null
        try {
            $fs = [System.IO.File]::OpenRead($Path)
            $buffer = New-Object byte[] 2
            [void]$fs.Read($buffer, 0, 2)
            $hexString = [System.BitConverter]::ToString($buffer)
            return $hexString -eq "50-4B"
        } finally {
            if ($fs) { $fs.Close() }
        }
    }

    # 1) Initial request
    Invoke-WebRequest -Uri $BaseUri -WebSession $Session -OutFile $TempPath -UseBasicParsing -TimeoutSec $requestTimeoutSec -ErrorAction Stop

    # 2) If already a zip, finish now
    if (Test-ZipSignature -Path $TempPath) {
        Move-Item -Path $TempPath -Destination $OutFilePath -Force
        return
    }

    Write-Host "  -> Intercepted Google Drive warning page. Attempting bypass..." -ForegroundColor DarkGray

    $htmlContent = Get-Content -Path $TempPath -Raw -ErrorAction SilentlyContinue

    # 3) Attempt to get confirm token from cookies first
    $token = $null
    try {
        $allCookies = @(
            $Session.Cookies.GetCookies([uri]"https://drive.google.com")
            $Session.Cookies.GetCookies([uri]"https://docs.google.com")
        )
        $cookieMatch = $allCookies | Where-Object { $_.Name -like "download_warning*" } | Select-Object -First 1
        if ($cookieMatch) { $token = $cookieMatch.Value }
    } catch { }

    # 4) Fallback token parsing from HTML
    if (-not $token) {
        if ($htmlContent -match 'confirm=([a-zA-Z0-9_-]+)') {
            $token = $Matches[1]
        }
        elseif ($htmlContent -match 'name="confirm"\s+value="([^"]+)"') {
            $token = $Matches[1]
        }
    }

    # 5) Build best possible follow-up URL from modern Google Drive warning page
    $FinalUri = $null
    if ($token) {
        $FinalUri = "https://drive.google.com/uc?export=download&id=$FileId&confirm=$token"
    }

    if (-not $FinalUri) {
        # Newer warning page pattern: action="https://drive.usercontent.google.com/download" + hidden inputs
        $actionMatch = [regex]::Match($htmlContent, 'action="(https://drive\.usercontent\.google\.com/download[^"]*)"')
        if ($actionMatch.Success) {
            $confirmVal = [regex]::Match($htmlContent, 'name="confirm"\s+value="([^"]+)"')
            $uuidVal = [regex]::Match($htmlContent, 'name="uuid"\s+value="([^"]+)"')

            $queryParts = @(
                "id=$([System.Uri]::EscapeDataString($FileId))",
                "export=download"
            )
            if ($confirmVal.Success) { $queryParts += "confirm=$([System.Uri]::EscapeDataString($confirmVal.Groups[1].Value))" }
            if ($uuidVal.Success) { $queryParts += "uuid=$([System.Uri]::EscapeDataString($uuidVal.Groups[1].Value))" }

            $FinalUri = "$($actionMatch.Groups[1].Value)?$($queryParts -join '&')"
        }
    }

    if (-not $FinalUri) {
        # Legacy fallback: direct /uc link in HTML
        $hrefMatch = [regex]::Match($htmlContent, 'href="(/uc\?export=download[^"]+)"')
        if ($hrefMatch.Success) {
            $decoded = $hrefMatch.Groups[1].Value.Replace('&amp;', '&')
            $FinalUri = "https://drive.google.com$decoded"
        }
    }

    # 6) Build a candidate list of URLs and try each until a valid ZIP is received
    $candidateUris = New-Object System.Collections.Generic.List[string]
    if ($FinalUri) { [void]$candidateUris.Add($FinalUri) }

    $uuidFromHtml = $null
    if ($htmlContent -match 'name="uuid"\s+value="([^"]+)"') {
        $uuidFromHtml = $Matches[1]
    }

    if ($token) {
        [void]$candidateUris.Add("https://drive.google.com/uc?export=download&id=$FileId&confirm=$token")
        [void]$candidateUris.Add("https://drive.usercontent.google.com/download?id=$FileId&export=download&confirm=$token")
    }

    [void]$candidateUris.Add("https://drive.google.com/uc?export=download&id=$FileId&confirm=t")
    [void]$candidateUris.Add("https://drive.usercontent.google.com/download?id=$FileId&export=download&confirm=t")
    if ($uuidFromHtml) {
        [void]$candidateUris.Add("https://drive.usercontent.google.com/download?id=$FileId&export=download&confirm=t&uuid=$([System.Uri]::EscapeDataString($uuidFromHtml))")
    }

    # Generic form action fallback: include only known-safe hidden inputs as query parameters
    $formActionMatch = [regex]::Match($htmlContent, '<form[^>]+action="([^"]+)"', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($formActionMatch.Success) {
        $formAction = $formActionMatch.Groups[1].Value
        if ($formAction.StartsWith('/')) { $formAction = "https://drive.google.com$formAction" }

        $hiddenInputRegex = '<input[^>]*type="hidden"[^>]*name="([^"]+)"[^>]*value="([^"]*)"[^>]*>'
        $hiddenInputs = [regex]::Matches($htmlContent, $hiddenInputRegex, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $allowedKeys = @('id','confirm','uuid','export','authuser')
        $queryParts = @()
        foreach ($input in $hiddenInputs) {
            $n = $input.Groups[1].Value
            $v = $input.Groups[2].Value
            if (-not [string]::IsNullOrWhiteSpace($n) -and ($allowedKeys -contains $n.ToLowerInvariant())) {
                $queryParts += "$([System.Uri]::EscapeDataString($n))=$([System.Uri]::EscapeDataString($v))"
            }
        }
        if ($queryParts.Count -gt 0) {
            $candidateFromForm = "$formAction?$($queryParts -join '&')"
            if ($candidateFromForm.Length -lt 2048) {
                $candidateUris.Add($candidateFromForm) | Out-Null
            }
        }
    }

    $tried = @()
    foreach ($uri in $candidateUris) {
        if ([string]::IsNullOrWhiteSpace($uri)) { continue }
        if ($tried -contains $uri) { continue }
        $tried += $uri

        try {
            Invoke-WebRequest -Uri $uri -WebSession $Session -OutFile $OutFilePath -UseBasicParsing -TimeoutSec $requestTimeoutSec -ErrorAction Stop
            if (Test-ZipSignature -Path $OutFilePath) {
                Remove-Item -Path $TempPath -Force -ErrorAction SilentlyContinue
                return
            }
            Remove-Item -Path $OutFilePath -Force -ErrorAction SilentlyContinue
        } catch {
            # Try next candidate URI
        }
    }

    Remove-Item -Path $TempPath -Force -ErrorAction SilentlyContinue
    throw "Downloaded file is not a valid ZIP archive (Google Drive returned HTML or a blocked response)."
}

# --- Helper Function to Resolve an Item ID from a Public Google Drive Folder ---
function Get-GDriveItemFromFolderByName {
    param(
        [Parameter(Mandatory=$true)][string]$FolderId,
        [Parameter(Mandatory=$true)][string]$ItemName
    )

    $embeddedUri = "https://drive.google.com/embeddedfolderview?id=$FolderId#list"
    $response = Invoke-WebRequest -Uri $embeddedUri -UseBasicParsing -ErrorAction Stop
    $html = $response.Content

    $anchorRegex = '<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>'
    $matches = [regex]::Matches($html, $anchorRegex, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    foreach ($m in $matches) {
        $href = $m.Groups[1].Value
        $nameRaw = $m.Groups[2].Value
        $nameClean = [System.Text.RegularExpressions.Regex]::Replace($nameRaw, '<.*?>', '')
        $nameClean = [System.Net.WebUtility]::HtmlDecode($nameClean).Trim()

        if ([string]::IsNullOrWhiteSpace($nameClean)) { continue }

        if ($nameClean -like "*$ItemName*") {
            $idMatch = [regex]::Match($href, '/file/d/([a-zA-Z0-9_-]+)')
            $itemType = "File"
            if (-not $idMatch.Success) {
                $idMatch = [regex]::Match($href, '/drive/folders/([a-zA-Z0-9_-]+)')
                $itemType = "Folder"
            }

            if ($idMatch.Success) {
                return [pscustomobject]@{
                    Name = $nameClean
                    Id   = $idMatch.Groups[1].Value
                    Type = $itemType
                }
            }
        }
    }

    return $null
}

function Test-ZipFile {
    param([Parameter(Mandatory=$true)][string]$Path)

    if (-not (Test-Path $Path)) { return $false }

    $fs = $null
    try {
        $fileInfo = Get-Item -Path $Path -ErrorAction Stop
        if ($fileInfo.Length -lt 2) { return $false }

        $fs = [System.IO.File]::OpenRead($Path)
        $buffer = New-Object byte[] 2
        [void]$fs.Read($buffer, 0, 2)
        $hexString = [System.BitConverter]::ToString($buffer)
        return ($hexString -eq "50-4B")
    } catch {
        return $false
    } finally {
        if ($fs) { $fs.Close() }
    }
}

# --- Global Directory Setup ---
$baseDir = Join-Path $env:USERPROFILE "Downloads"
$sharedDriveFolderId = "1_o7vHTNPyofzaiPgXAyeZg_jxRGeKG8R"
Write-Host "`n--- Setting up Base Directory ---" -ForegroundColor Cyan
Write-Host "Using base directory at: $baseDir" -ForegroundColor Green

do {
    $currentHostName = $env:COMPUTERNAME

    Write-Host "`n--- Task Selection ---" -ForegroundColor Cyan
    Write-Host "[1] Hostname Configuration"
    Write-Host "[2] AV/VPN/Bloatware Scan & Uninstall"
    Write-Host "[3] HP Wolf Security Uninstaller"
    Write-Host "[4] Trend Micro Uninstaller"
    Write-Host "[5] CrowdStrike Falcon Installer"
    Write-Host "[A] Run All"
    Write-Host "[Q] Quit"

    $selectionInput = Read-Host "Select task numbers (comma-separated), A for all, or Q to quit"
    if ($selectionInput -match '^[Qq]$') { break }

    $selectedParts = @()

    if ($selectionInput -match '^[Aa]$') {
        $selectedParts = 1..5
    } else {
        $selectedParts = $selectionInput -split ',' |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -match '^[1-5]$' } |
            ForEach-Object { [int]$_ } |
            Select-Object -Unique
    }

    if (-not $selectedParts -or $selectedParts.Count -eq 0) {
        Write-Host "No valid tasks selected. Please try again." -ForegroundColor Yellow
        continue
    }

# PART 1: Change Hostname
if ($selectedParts -contains 1) {
    Write-Host "`n--- Hostname Configuration ---" -ForegroundColor Cyan
    Write-Host "Current Hostname: $currentHostName" -ForegroundColor Yellow

    $newHostName = Read-Host "Enter new Hostname (or press Enter to skip)"

    if (![string]::IsNullOrWhiteSpace($newHostName)) {
        if ($newHostName -eq $currentHostName) {
            Write-Host "New hostname is the same as current. Skipping rename." -ForegroundColor Yellow
        } else {
            try {
                Rename-Computer -NewName $newHostName -Force -ErrorAction Stop
                Write-Host "Hostname successfully changed to: $newHostName" -ForegroundColor Green
                Write-Host "NOTE: A reboot is required for the name change to fully take effect." -ForegroundColor Magenta
            } catch {
                Write-Error "Failed to change hostname: $($_.Exception.Message)"
            }
        }
    } else {
        Write-Host "Skipping hostname change." -ForegroundColor Yellow
    }
}

# PART 2: AV / VPN / Bloatware Cross-Reference & Uninstaller
if ($selectedParts -contains 2) {
Write-Host "`n--- Loading Installed Applications ---" -ForegroundColor Cyan

$blocklistUrl = "https://raw.githubusercontent.com/Justine25-debug/QCLGU/refs/heads/main/Uninstall_List.txt"

 $blockKeywords = @()
try {
    Write-Host "Downloading latest uninstall blocklist from remote source..." -ForegroundColor Yellow
    $blocklistResponse = Invoke-WebRequest -Uri $blocklistUrl -UseBasicParsing -TimeoutSec 60 -ErrorAction Stop
    $blockKeywords = ($blocklistResponse.Content -split "`r?`n") | Where-Object { $_.Trim() -ne "" }
    Write-Host "Loaded $($blockKeywords.Count) blocklist entries from remote source." -ForegroundColor Green
} catch {
    Write-Warning "Could not download remote blocklist. Continuing with an empty blocklist. Error: $($_.Exception.Message)"
}

$registryPaths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

$allApps = Get-ItemProperty $registryPaths -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -and $_.UninstallString } |
    Select-Object DisplayName, DisplayVersion, Publisher, UninstallString |
    Sort-Object DisplayName -Unique

if ($allApps.Count -eq 0) {
    Write-Host "No applications found on this system." -ForegroundColor Yellow
} else {
    $flaggedApps = @()
    foreach ($app in $allApps) {
        foreach ($keyword in $blockKeywords) {
            if ($app.DisplayName -match [regex]::Escape($keyword.Trim())) {
                $flaggedApps += $app
                break
            }
        }
    }

    function Start-UninstallLoop ($appList, $listName) {
        while ($true) {
            Write-Host "`n--- $listName ---`n" -ForegroundColor Green

            if ($listName -eq "All Installed Applications") {
                $columns = 3
                $indexWidth = (("[{0}]" -f $appList.Count).Length + 1)
                $maxNameWidth = ($appList | ForEach-Object { $_.DisplayName.Length } | Measure-Object -Maximum).Maximum
                $nameWidth = [Math]::Min([Math]::Max($maxNameWidth, 20), 34)
                $columnWidth = $indexWidth + $nameWidth + 3
                $rows = [Math]::Ceiling($appList.Count / [double]$columns)

                for ($row = 0; $row -lt $rows; $row++) {
                    $line = ""
                    for ($col = 0; $col -lt $columns; $col++) {
                        $idx = $row + ($col * $rows) + 1
                        if ($idx -le $appList.Count) {
                            $number = ("[{0}]" -f $idx).PadRight($indexWidth)
                            $name = $appList[$idx - 1].DisplayName
                            if ($name.Length -gt $nameWidth) {
                                $name = $name.Substring(0, $nameWidth - 3) + "..."
                            }
                            $entry = "$number $name"
                            $line += $entry.PadRight($columnWidth)
                        }
                    }
                    Write-Host $line
                }
            } else {
                $i = 1
                foreach ($app in $appList) {
                    $number = "{0,-4}" -f "[$i]"
                    Write-Host "$number $($app.DisplayName) (Version: $($app.DisplayVersion))"
                    $i++
                }
            }

            Write-Host ""
            $allowSelectAll = ($listName -ne "All Installed Applications")
            if ($allowSelectAll) {
                $selection = Read-Host "Enter a number, comma-separated numbers (e.g., 1,2,3), 'A' for all, or press Enter to skip"
            } else {
                $selection = Read-Host "Enter a number, comma-separated numbers (e.g., 1,2,3), or press Enter to skip"
            }

            if ([string]::IsNullOrWhiteSpace($selection)) { break }

            if ($allowSelectAll -and $selection -match '^[Aa]$') {
                $confirmAll = Read-Host "Are you sure you want to uninstall ALL listed apps? (Y/N)"
                if ($confirmAll -match '^[Yy]$') {
                    foreach ($targetApp in $appList) {
                        Write-Host "`nUninstalling: $($targetApp.DisplayName)" -ForegroundColor Magenta
                        $uninstallString = $targetApp.UninstallString
                        Write-Host "Executing: $uninstallString" -ForegroundColor Cyan
                        try {
                            Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$uninstallString`"" -Wait -NoNewWindow
                            Write-Host "Finished." -ForegroundColor Green
                            
                            # Web blocklist only 
                        } catch {
                            Write-Error "Failed to launch uninstaller for $($targetApp.DisplayName)."
                        }
                    }
                    break 
                } else {
                    Write-Host "Bulk uninstallation cancelled." -ForegroundColor Yellow
                }
            }
            elseif ($selection -match '^[\d\s,]+$') {
                $selectedIndexes = $selection -split ',' | ForEach-Object { $_.Trim() } | 
                    Where-Object { $_ -match '^\d+$' -and [int]$_ -ge 1 -and [int]$_ -le $appList.Count } | Select-Object -Unique

                if ($selectedIndexes.Count -eq 0) {
                    Write-Warning "No valid application numbers selected. Please try again."
                    continue
                }

                Write-Host "`nYou selected $($selectedIndexes.Count) application(s) for removal." -ForegroundColor Magenta
                $confirm = Read-Host "Are you sure you want to uninstall the selected app(s)? (Y/N)"

                if ($confirm -match '^[Yy]$') {
                    foreach ($idx in $selectedIndexes) {
                        $selectedApp = $appList[[int]$idx - 1] 
                        Write-Host "`nUninstalling: $($selectedApp.DisplayName)" -ForegroundColor Magenta
                        
                        $uninstallString = $selectedApp.UninstallString
                        Write-Host "Executing: $uninstallString" -ForegroundColor Cyan
                        try {
                            Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$uninstallString`"" -Wait -NoNewWindow
                            Write-Host "Uninstallation process finished/closed." -ForegroundColor Green
                            
                            # Web blocklist only
                        } catch {
                            Write-Error "Failed to launch uninstaller for $($selectedApp.DisplayName)."
                        }
                    }
                } else {
                    Write-Host "Uninstallation cancelled." -ForegroundColor Yellow
                }
            } else {
                Write-Warning "Invalid selection. Please enter numbers separated by commas, or 'A'."
            }
        }
    }

    if ($flaggedApps.Count -gt 0) {
        Write-Host "`nWARNING: Found apps matching the online blocklist!" -ForegroundColor Red
        Start-UninstallLoop -appList $flaggedApps -listName "Flagged AV/VPN/Bloatware"
    } else {
        Write-Host "Clean! No known AV/VPN/Bloatware from the online blocklist were found." -ForegroundColor Green
    }

    Write-Host ""
    $showFull = Read-Host "Do you want to view the FULL list of installed applications? (Y/N)"
    if ($showFull -match '^[Yy]$') {
        Start-UninstallLoop -appList $allApps -listName "All Installed Applications"
    }
}
}

# PART 3: HP Wolf Security Uninstaller (WMIC-Free Edition)
if ($selectedParts -contains 3) {
    Write-Host "`n--- HP Wolf Security Uninstaller Section ---" -ForegroundColor Cyan
    Write-Host "Attempting to uninstall HP Wolf Security components in required order..." -ForegroundColor Magenta

    $hpComponents = @(
        "HP Wolf Security",
        "HP Wolf Security - Console",
        "HP Security Update Service"
    )

    foreach ($component in $hpComponents) {
        Write-Host "`nSearching for: $component..." -ForegroundColor Yellow
        $uninstalled = $false

        $pkg = Get-Package -Name $component -ErrorAction SilentlyContinue
        if ($pkg -and -not $uninstalled) {
            Write-Host "Found $($pkg.Name). Uninstalling via Get-Package..." -ForegroundColor Cyan
            try {
                $pkg | Uninstall-Package -Force -ErrorAction Stop
                Write-Host "Finished processing $component." -ForegroundColor Green
                $uninstalled = $true
            } catch {
                Write-Warning "Get-Package uninstall failed. Trying fallback..."
            }
        }

        if (-not $uninstalled) {
            $cimPkg = Get-CimInstance -ClassName Win32_Product -Filter "Name = '$component'" -ErrorAction SilentlyContinue
            if ($cimPkg) {
                Write-Host "Found $($cimPkg.Name). Uninstalling via WMI/CIM..." -ForegroundColor Cyan
                try {
                    Invoke-CimMethod -InputObject $cimPkg -MethodName Uninstall | Out-Null
                    Write-Host "Finished processing $component." -ForegroundColor Green
                    $uninstalled = $true
                } catch {
                    Write-Warning "CIM uninstall failed. Trying registry fallback..."
                }
            }
        }

        if (-not $uninstalled) {
            $regPath = @("HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
            $regApp = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -eq $component }
            
            if ($regApp) {
                $guid = $regApp.PSChildName
                if ($guid -match "^\{.*\}$") {
                    Write-Host "Found Registry GUID ($guid). Uninstalling via msiexec..." -ForegroundColor Cyan
                    try {
                        Start-Process -FilePath "msiexec.exe" -ArgumentList "/x `"$guid`" /qn /norestart" -Wait -NoNewWindow
                        Write-Host "Finished processing $component." -ForegroundColor Green
                        $uninstalled = $true
                     } catch {
                        Write-Error "Failed to execute msiexec for $component."
                    }
                }
            }
        }
        
        if (-not $uninstalled) {
            Write-Host "$component was not found on this system or is already removed." -ForegroundColor DarkGray
        }
    }
    
    Write-Host "`nHP Wolf Security uninstallation sequence complete. A reboot may be required." -ForegroundColor Yellow
}

# PART 4: Trend Micro Custom Uninstaller (Web Download)
if ($selectedParts -contains 4) {
    Write-Host "`n--- Trend Micro Uninstaller Section ---" -ForegroundColor Cyan
    $tmGithubAssetUrl = "https://github.com/Justine25-debug/QCLGU/releases/download/v1.0/TM.Uninstall.-.April.08.2026.zip"
    $tmZipPath = Join-Path $baseDir "TM_Uninstaller.zip"
    $tmExtractPath = Join-Path $baseDir "TMUNINSTALL"
    $tmExeName = "V1ESUninstallTool.exe"
    
    try {
        $tmTargetExeInfo = Get-ChildItem -Path $tmExtractPath -Filter $tmExeName -Recurse | Select-Object -First 1

        if ($tmTargetExeInfo) {
            Write-Host "Existing Trend Micro extracted files found. Skipping download/extraction." -ForegroundColor Green
            $tmTargetExe = $tmTargetExeInfo.FullName
            Write-Host "Launching Trend Micro Uninstaller from $tmTargetExe..." -ForegroundColor Magenta
            Set-Location -Path (Split-Path $tmTargetExe -Parent)
            
            # Launches the uninstaller
            Start-Process -FilePath $tmTargetExe
            
            Write-Host "`n[NOTICE] The Trend Micro Uninstaller has been launched in a separate window." -ForegroundColor Yellow
            Write-Host "Please complete the uninstallation steps in that window." -ForegroundColor Gray
        } else {
            Write-Host "Downloading Trend Micro package from GitHub Release..." -ForegroundColor Magenta

            $isZip = $false
            if (Test-Path $tmZipPath) {
                if (Test-ZipFile -Path $tmZipPath) {
                    Write-Host "Existing Trend Micro ZIP found at $tmZipPath. Skipping download." -ForegroundColor Green
                    $isZip = $true
                } else {
                    Write-Warning "Existing TM_Uninstaller.zip is invalid. Re-downloading..."
                    Remove-Item -Path $tmZipPath -Force -ErrorAction SilentlyContinue
                }
            }

            if (-not $isZip) {
                Invoke-WebRequest -Uri $tmGithubAssetUrl -OutFile $tmZipPath -UseBasicParsing -TimeoutSec 300 -Headers @{ "User-Agent" = "Mozilla/5.0" } -ErrorAction Stop
                $isZip = Test-ZipFile -Path $tmZipPath
            }

            if (-not $isZip) {
                Write-Warning "Direct GitHub request did not return a ZIP. Opening browser for manual download from release page."
                Start-Process "https://github.com/Justine25-debug/QCLGU/releases/tag/v1.0"
                Read-Host "Download the Trend Micro ZIP from the opened page, then press Enter to continue"

                $downloadedTmZip = Get-ChildItem -Path $baseDir -Filter "*TM*Uninstall*.zip" -File -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 1

                if (-not $downloadedTmZip) {
                    $downloadedTmZip = Get-ChildItem -Path $baseDir -Filter "*.zip" -File -ErrorAction SilentlyContinue |
                        Sort-Object LastWriteTime -Descending |
                        Select-Object -First 1
                }

                if (-not $downloadedTmZip) {
                    throw "No ZIP file found in Downloads after manual GitHub step."
                }

                Copy-Item -Path $downloadedTmZip.FullName -Destination $tmZipPath -Force
                Write-Host "Using manually downloaded ZIP: $($downloadedTmZip.FullName)" -ForegroundColor Green
            }

            Write-Host "Download complete. Extracting $tmZipPath..." -ForegroundColor Green
            Expand-Archive -Path $tmZipPath -DestinationPath $tmExtractPath -Force

            $tmTargetExeInfo = Get-ChildItem -Path $tmExtractPath -Filter $tmExeName -Recurse | Select-Object -First 1
            if ($tmTargetExeInfo) {
                $tmTargetExe = $tmTargetExeInfo.FullName
                Write-Host "Launching Trend Micro Uninstaller from $tmTargetExe..." -ForegroundColor Magenta
                Set-Location -Path (Split-Path $tmTargetExe -Parent)
                Start-Process -FilePath $tmTargetExe

                Write-Host "`n[NOTICE] The Trend Micro Uninstaller has been launched in a separate window." -ForegroundColor Yellow
                Write-Host "Please complete the uninstallation steps in that window." -ForegroundColor Gray
            } else {
                Write-Host "Error: Could not find $tmExeName within the downloaded archive." -ForegroundColor Red
            }
        }
    } catch {
        $tmRecovered = $false
        $tmTargetExeInfo = Get-ChildItem -Path $tmExtractPath -Filter $tmExeName -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($tmTargetExeInfo) {
            Write-Warning "Could not use archive, but existing extracted Trend Micro files were found. Launching existing executable."
            $tmTargetExe = $tmTargetExeInfo.FullName
            Write-Host "Launching Trend Micro Uninstaller from $tmTargetExe..." -ForegroundColor Magenta
            Set-Location -Path (Split-Path $tmTargetExe -Parent)
            Start-Process -FilePath $tmTargetExe
            $tmRecovered = $true
        }

        if (-not $tmRecovered) {
            Write-Warning "Failed to prepare Trend Micro Uninstaller. Error: $($_.Exception.Message)"
            $manualTmZip = Read-Host "Enter local path to Trend Micro ZIP (or press Enter to skip)"
        }

        if (-not $tmRecovered -and -not [string]::IsNullOrWhiteSpace($manualTmZip) -and (Test-Path $manualTmZip)) {
            try {
                Copy-Item -Path $manualTmZip -Destination $tmZipPath -Force
                Write-Host "Using local ZIP. Extracting $tmZipPath..." -ForegroundColor Green
                Expand-Archive -Path $tmZipPath -DestinationPath $tmExtractPath -Force

                $tmTargetExeInfo = Get-ChildItem -Path $tmExtractPath -Filter $tmExeName -Recurse | Select-Object -First 1
                if ($tmTargetExeInfo) {
                    $tmTargetExe = $tmTargetExeInfo.FullName
                    Write-Host "Launching Trend Micro Uninstaller from $tmTargetExe..." -ForegroundColor Magenta
                    Set-Location -Path (Split-Path $tmTargetExe -Parent)
                    Start-Process -FilePath $tmTargetExe

                    Write-Host "`n[NOTICE] The Trend Micro Uninstaller has been launched in a separate window." -ForegroundColor Yellow
                    Write-Host "Please complete the uninstallation steps in that window." -ForegroundColor Gray
                } else {
                    Write-Error "Could not find $tmExeName in the local ZIP archive."
                }
            } catch {
                Write-Error "Failed to process local Trend Micro ZIP. Error: $_"
            }
        } elseif (-not $tmRecovered) {
            Write-Host "Skipping Trend Micro Uninstaller." -ForegroundColor Yellow
        }
    }
}

# PART 5: CrowdStrike Falcon Installer (Web Download)
if ($selectedParts -contains 5) {
    Write-Host "`n--- CrowdStrike Falcon Installer Section ---" -ForegroundColor Cyan
    $csGithubAssetUrl = "https://github.com/Justine25-debug/QCLGU/releases/download/v1.0/CS.Falcon.zip"
    $csZipPath = Join-Path $baseDir "CS_Installer.zip"
    $csExtractPath = Join-Path $baseDir "CSFalcon"
    $csExeName = "QCLGU-CS-installer.exe"
    
    try {
        $csTargetExeInfo = Get-ChildItem -Path $csExtractPath -Filter $csExeName -Recurse | Select-Object -First 1

        if ($csTargetExeInfo) {
            Write-Host "Existing CrowdStrike extracted files found. Skipping download/extraction." -ForegroundColor Green
            $csTargetExe = $csTargetExeInfo.FullName
            Write-Host "Launching CrowdStrike Installer from $csTargetExe..." -ForegroundColor Magenta
            Set-Location -Path (Split-Path $csTargetExe -Parent)
            
            # Launches the installer
            Start-Process -FilePath $csTargetExe
            
            Write-Host "`n[NOTICE] The CrowdStrike Installer has been launched in a separate window." -ForegroundColor Yellow
            Write-Host "Please complete the installation steps in that window." -ForegroundColor Gray
        } else {
            Write-Host "Downloading CrowdStrike package from GitHub Release..." -ForegroundColor Magenta

            $isZip = $false
            if (Test-Path $csZipPath) {
                if (Test-ZipFile -Path $csZipPath) {
                    Write-Host "Existing CrowdStrike ZIP found at $csZipPath. Skipping download." -ForegroundColor Green
                    $isZip = $true
                } else {
                    Write-Warning "Existing CS_Installer.zip is invalid. Re-downloading..."
                    Remove-Item -Path $csZipPath -Force -ErrorAction SilentlyContinue
                }
            }

            if (-not $isZip) {
                Invoke-WebRequest -Uri $csGithubAssetUrl -OutFile $csZipPath -UseBasicParsing -TimeoutSec 180 -Headers @{ "User-Agent" = "Mozilla/5.0" } -ErrorAction Stop
                $isZip = Test-ZipFile -Path $csZipPath
            }

            if (-not $isZip) {
                Write-Warning "Direct GitHub request did not return a ZIP. Opening browser for manual download from release page."
                Start-Process "https://github.com/Justine25-debug/QCLGU/releases/tag/v1.0"
                Read-Host "Download the CrowdStrike ZIP from the opened page, then press Enter to continue"

                $downloadedCsZip = Get-ChildItem -Path $baseDir -Filter "*CS*Falcon*.zip" -File -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First 1

                if (-not $downloadedCsZip) {
                    $downloadedCsZip = Get-ChildItem -Path $baseDir -Filter "*.zip" -File -ErrorAction SilentlyContinue |
                        Sort-Object LastWriteTime -Descending |
                        Select-Object -First 1
                }

                if (-not $downloadedCsZip) {
                    throw "No ZIP file found in Downloads after manual GitHub step."
                }

                Copy-Item -Path $downloadedCsZip.FullName -Destination $csZipPath -Force
                Write-Host "Using manually downloaded ZIP: $($downloadedCsZip.FullName)" -ForegroundColor Green
            }

            Write-Host "Download complete. Extracting $csZipPath..." -ForegroundColor Green
            Expand-Archive -Path $csZipPath -DestinationPath $csExtractPath -Force

            $csTargetExeInfo = Get-ChildItem -Path $csExtractPath -Filter $csExeName -Recurse | Select-Object -First 1
            if ($csTargetExeInfo) {
                $csTargetExe = $csTargetExeInfo.FullName
                Write-Host "Launching CrowdStrike Installer from $csTargetExe..." -ForegroundColor Magenta
                Set-Location -Path (Split-Path $csTargetExe -Parent)
                Start-Process -FilePath $csTargetExe

                Write-Host "`n[NOTICE] The CrowdStrike Installer has been launched in a separate window." -ForegroundColor Yellow
                Write-Host "Please complete the installation steps in that window." -ForegroundColor Gray
            } else {
                Write-Host "Error: Could not find $csExeName within the downloaded archive." -ForegroundColor Red
            }
        }
    } catch {
        $csTargetExeInfo = Get-ChildItem -Path $csExtractPath -Filter $csExeName -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($csTargetExeInfo) {
            Write-Warning "Could not use archive, but existing extracted CrowdStrike files were found. Launching existing executable."
            $csTargetExe = $csTargetExeInfo.FullName
            Write-Host "Launching CrowdStrike Installer from $csTargetExe..." -ForegroundColor Magenta
            Set-Location -Path (Split-Path $csTargetExe -Parent)
            Start-Process -FilePath $csTargetExe
        } else {
            Write-Error "Failed to prepare or execute CrowdStrike Installer. Error: $_"
        }
    }
}

Write-Host "`n--- All Tasks Completed ---" -ForegroundColor Cyan

    $runAgain = Read-Host "Return to task menu? (Y/N)"
    if ($runAgain -notmatch '^[Yy]$') { break }
} while ($true)

Write-Host "Exiting automation menu." -ForegroundColor Cyan
