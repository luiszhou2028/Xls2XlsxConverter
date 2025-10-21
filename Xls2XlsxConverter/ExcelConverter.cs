using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Xls2XlsxConverter
{
    public class ExcelConverter : IDisposable
    {
        private Application _excelApp;
        private bool _disposed = false;
        private bool _cancellationRequested = false;

        public ExcelConverter()
        {
            InitializeExcelApplication();
        }

        private void InitializeExcelApplication()
        {
            try
            {
                _excelApp = new Application
                {
                    Visible = false,
                    DisplayAlerts = false,
                    ScreenUpdating = false,
                    EnableEvents = false,
                    AskToUpdateLinks = false,
                    AlertBeforeOverwriting = false
                };
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException("无法启动Excel应用程序。请确保已安装Microsoft Excel。", ex);
            }
        }

        public void ConvertXlsToXlsx(string xlsFilePath, string xlsxFilePath)
        {
            if (_cancellationRequested)
                throw new OperationCanceledException("转换操作已被取消");

            if (!File.Exists(xlsFilePath))
                throw new FileNotFoundException($"源文件不存在: {xlsFilePath}");

            Workbook workbook = null;
            try
            {
                // 打开XLS文件
                workbook = _excelApp.Workbooks.Open(
                    Filename: xlsFilePath,
                    UpdateLinks: 0,
                    ReadOnly: true,
                    Format: 5,
                    Password: Type.Missing,
                    WriteResPassword: Type.Missing,
                    IgnoreReadOnlyRecommended: true,
                    Origin: Type.Missing,
                    Delimiter: Type.Missing,
                    Editable: false,
                    Notify: false,
                    Converter: Type.Missing,
                    AddToMru: false,
                    Local: Type.Missing,
                    CorruptLoad: Type.Missing
                );

                if (_cancellationRequested)
                    throw new OperationCanceledException("转换操作已被取消");

                // 删除已存在的目标文件
                if (File.Exists(xlsxFilePath))
                {
                    File.Delete(xlsxFilePath);
                }

                // 保存为XLSX格式
                workbook.SaveAs(
                    Filename: xlsxFilePath,
                    FileFormat: XlFileFormat.xlOpenXMLWorkbook,
                    Password: Type.Missing,
                    WriteResPassword: Type.Missing,
                    ReadOnlyRecommended: false,
                    CreateBackup: false,
                    AccessMode: XlSaveAsAccessMode.xlNoChange,
                    ConflictResolution: XlSaveConflictResolution.xlLocalSessionChanges,
                    AddToMru: false,
                    TextCodepage: Type.Missing,
                    TextVisualLayout: Type.Missing,
                    Local: Type.Missing
                );
            }
            catch (COMException ex)
            {
                throw new InvalidOperationException($"转换文件时发生错误: {ex.Message}", ex);
            }
            finally
            {
                // 关闭工作簿
                if (workbook != null)
                {
                    try
                    {
                        workbook.Close(SaveChanges: false);
                        Marshal.ReleaseComObject(workbook);
                    }
                    catch (COMException)
                    {
                        // 忽略关闭时的错误
                    }
                }
            }
        }

        public void CancelConversion()
        {
            _cancellationRequested = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _cancellationRequested = true;
                }

                if (_excelApp != null)
                {
                    try
                    {
                        _excelApp.Quit();
                        Marshal.ReleaseComObject(_excelApp);
                    }
                    catch (COMException)
                    {
                        // 忽略退出时的错误
                    }
                    finally
                    {
                        _excelApp = null;
                    }
                }

                _disposed = true;
            }
        }

        ~ExcelConverter()
        {
            Dispose(false);
        }
    }

    public class ConversionOptions
    {
        public string InputFolder { get; set; }
        public string OutputFolder { get; set; }
        public bool IncludeSubfolders { get; set; } = true;
        public bool OverwriteExisting { get; set; } = false;
    }
}