@echo off
chcp 65001 >nul
echo.
echo ===============================================
echo           XLS到XLSX批量转换器 v1.0
echo ===============================================
echo.
echo 正在启动程序...
echo.

cd /d "%~dp0"

if exist "Xls2XlsxConverter\bin\Release\net6.0-windows\Xls2XlsxConverter.exe" (
    start "" "Xls2XlsxConverter\bin\Release\net6.0-windows\Xls2XlsxConverter.exe"
    echo 程序启动成功！
) else (
    echo 错误：找不到可执行文件！
    echo 请先运行 "编译程序.bat" 来编译项目。
    echo.
    pause
)

echo.
echo 如需重新编译程序，请运行 "编译程序.bat"
echo.
timeout /t 3 >nul