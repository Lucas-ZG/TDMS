@rem Lucas.ZhuGe
@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion

:: 先固定目前環境的 Python 絕對路徑，再縮小打包時的 DLL 搜尋 PATH。
:: 避免外部工具（例如 Poppler）同名 DLL 被 PyInstaller 誤收進成品。
for %%I in (python.exe) do set "PYTHON_EXE=%%~$PATH:I"

if not defined PYTHON_EXE (
    echo ❌ 找不到 python。請先安裝或啟用正確的 Python 環境。
    pause
    exit /b
)

echo 🔍 搜尋當前資料夾內的 .py 檔案...

set "count=0"
set "lastpy="

for %%f in (*.py) do (
    set /a count+=1
    set "lastpy=%%~ff"
)

if %count%==0 (
    echo ❌ 找不到任何 .py 檔案，請確認資料夾內是否有 Python 程式。
    pause
    exit /b
)

:: 檢查目前 Python 環境內的 PyInstaller 是否可用
"%PYTHON_EXE%" -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 找不到 pyinstaller。請先：pip install pyinstaller
    pause
    exit /b
)

:: 臨時清單檔：存放使用者選到的檔案（每行一個完整路徑）
set "SELFILE=%TEMP%\selected_py_%RANDOM%.txt"
if exist "%SELFILE%" del /f /q "%SELFILE%" >nul 2>&1

if %count% GTR 1 (
    echo ⚠️ 發現多個 .py 檔案，請選擇要打包的檔案（可多選）...

    :: 先嘗試使用 WPF 的 OpenFileDialog（DPI 感知，顯示較清晰）
    powershell -NoProfile -Command ^
      "try { " ^
      "  Add-Type -AssemblyName PresentationFramework | Out-Null; " ^
      "  $dlg = New-Object Microsoft.Win32.OpenFileDialog; " ^
      "  $dlg.InitialDirectory = (Get-Location).Path; " ^
      "  $dlg.Filter = 'Python (*.py)|*.py'; " ^
      "  $dlg.Multiselect = $true; " ^
      "  $dlg.Title = '選擇要打包的 .py 檔案（可多選）'; " ^
      "  if ($dlg.ShowDialog()) { $dlg.FileNames } " ^
      "} catch { " ^
      "  Add-Type -AssemblyName System.Windows.Forms | Out-Null; " ^
      "  $fd = New-Object System.Windows.Forms.OpenFileDialog; " ^
      "  $fd.InitialDirectory = (Get-Location).Path; " ^
      "  $fd.Filter = 'Python (*.py)|*.py'; " ^
      "  $fd.Multiselect = $true; " ^
      "  $fd.Title = '選擇要打包的 .py 檔案（可多選）'; " ^
      "  if ($fd.ShowDialog() -eq 'OK') { $fd.FileNames } " ^
      "}" > "%SELFILE%"

    if not exist "%SELFILE%" (
        echo ❌ 未選擇任何檔案。
        pause
        exit /b
    )

    set "picked=0"
    for /f "usebackq delims=" %%A in ("%SELFILE%") do (
        if not "%%~A"=="" (
            set /a picked+=1
            call :BUILD_ONE "%%~fA"
        )
    )
    if !picked! EQU 0 (
        echo ❌ 未選擇任何檔案。
        del /f /q "%SELFILE%" >nul 2>&1
        pause
        exit /b
    )
) else (
    echo ✅ 準備打包: "%lastpy%"
    call :BUILD_ONE "%lastpy%"
)

if exist "%SELFILE%" del /f /q "%SELFILE%" >nul 2>&1
echo.
echo ✅ 全部打包完成。
pause
exit /b


:BUILD_ONE
:: %~1 = 要打包的 .py 完整路徑
set "PYFILE=%~1"
set "BASENAME=%~n1"
echo.
echo ===============================
echo 🚧 正在打包: "%PYFILE%"
echo ===============================

:: 針對單一檔案清理舊產物
rmdir /s /q build >nul 2>&1
rmdir /s /q dist >nul 2>&1
del /f /q "%BASENAME%.spec" >nul 2>&1

:: 執行打包（無主控台、單檔）。只保留 Windows 系統 DLL 搜尋路徑，
:: Python 與套件由上方固定的直譯器絕對路徑提供。
set "ORIGINAL_PATH=%PATH%"
set "PATH=%SystemRoot%\System32;%SystemRoot%;%SystemRoot%\System32\Wbem"
"%PYTHON_EXE%" -m PyInstaller --clean --noconfirm --noconsole --onefile "%PYFILE%"
set "BUILD_EXIT=!ERRORLEVEL!"
set "PATH=%ORIGINAL_PATH%"
if not "!BUILD_EXIT!"=="0" (
    echo ❌ 打包失敗: "%PYFILE%"
    goto :EOF
)

:: 移動輸出到目前資料夾並改成與 .py 同名
if exist "dist\%BASENAME%.exe" (
    move /y "dist\%BASENAME%.exe" ".\%BASENAME%.exe" >nul
    echo ✅ 完成: "%BASENAME%.exe"
) else (
    set "foundexe="
    for %%E in (dist\*.exe) do set "foundexe=%%~fE"
    if defined foundexe (
        move /y "!foundexe!" ".\%BASENAME%.exe" >nul
        echo ✅ 完成（重新命名）: "%BASENAME%.exe"
    ) else (
        echo ⚠️ 未找到輸出 exe，請檢查 PyInstaller 設定。
    )
)

:: 清除暫存
rmdir /s /q build >nul 2>&1
rmdir /s /q dist >nul 2>&1
del /f /q "%BASENAME%.spec" >nul 2>&1
goto :EOF
