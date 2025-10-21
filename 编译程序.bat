@echo off
chcp 65001 >nul
echo.
echo ===============================================
echo           编译XLS到XLSX转换器
echo ===============================================
echo.

cd /d "%~dp0"

echo 正在检查.NET 6.0环境...
dotnet --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未找到.NET 6.0 SDK！
    echo 请先安装.NET 6.0 SDK：https://dotnet.microsoft.com/download/dotnet/6.0
    echo.
    pause
    exit /b 1
)

echo 找到.NET SDK版本：
dotnet --version
echo.

echo 正在还原NuGet包...
dotnet restore
if errorlevel 1 (
    echo 错误：包还原失败！
    pause
    exit /b 1
)

echo.
echo 正在编译项目（Release模式）...
dotnet build -c Release
if errorlevel 1 (
    echo 错误：编译失败！
    pause
    exit /b 1
)

echo.
echo ===============================================
echo              编译成功完成！
echo ===============================================
echo.
echo 可执行文件位置：
echo Xls2XlsxConverter\bin\Release\net6.0-windows\Xls2XlsxConverter.exe
echo.
echo 现在可以运行 "运行程序.bat" 来启动程序
echo.
pause