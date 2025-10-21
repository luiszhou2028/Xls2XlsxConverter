using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Xls2XlsxConverter
{
    public partial class MainForm : Form
    {
        private readonly ExcelConverter _converter;
        private bool _isConverting = false;
        private NotifyIcon _notifyIcon;
        private ContextMenuStrip _trayMenu;
        private bool _hasShownTrayTip = false;
        private FileSystemWatcher _fileWatcher;
        private readonly ConcurrentQueue<string> _pendingFiles = new ConcurrentQueue<string>();
        private readonly ConcurrentDictionary<string, byte> _queuedFiles = new ConcurrentDictionary<string, byte>();
        private readonly ConcurrentDictionary<string, int> _autoRetryCounts = new ConcurrentDictionary<string, int>();
        private CancellationTokenSource _watcherCancellation;
        private Task _backgroundProcessorTask;
        private ConversionOptions _autoModeOptions;
        private bool _isAutoMonitoring = false;
        private bool _suppressAutoMonitorEvent = false;
        private readonly SemaphoreSlim _conversionSemaphore = new SemaphoreSlim(1, 1);

        public MainForm()
        {
            InitializeComponent();
            _converter = new ExcelConverter();
            InitializeTrayIcon();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // 窗体设置
            this.Text = "XLS到XLSX批量转换器 v1.0";
            this.Size = new Size(650, 500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = true;

            // 创建控件
            CreateControls();

            this.Resize += MainForm_Resize;
            
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void CreateControls()
        {
            // 输入文件夹选择
            var lblInputFolder = new Label
            {
                Text = "输入文件夹（包含XLS文件）:",
                Location = new Point(20, 20),
                Size = new Size(200, 20),
                Font = new Font("Microsoft YaHei", 9F)
            };
            this.Controls.Add(lblInputFolder);

            var txtInputFolder = new TextBox
            {
                Name = "txtInputFolder",
                Location = new Point(20, 45),
                Size = new Size(450, 25),
                Font = new Font("Microsoft YaHei", 9F),
                ReadOnly = true
            };
            this.Controls.Add(txtInputFolder);

            var btnSelectInput = new Button
            {
                Name = "btnSelectInput",
                Text = "浏览...",
                Location = new Point(480, 44),
                Size = new Size(80, 27),
                Font = new Font("Microsoft YaHei", 9F),
                UseVisualStyleBackColor = true
            };
            btnSelectInput.Click += BtnSelectInput_Click;
            this.Controls.Add(btnSelectInput);

            // 输出文件夹选择
            var lblOutputFolder = new Label
            {
                Text = "输出文件夹（保存XLSX文件）:",
                Location = new Point(20, 85),
                Size = new Size(200, 20),
                Font = new Font("Microsoft YaHei", 9F)
            };
            this.Controls.Add(lblOutputFolder);

            var txtOutputFolder = new TextBox
            {
                Name = "txtOutputFolder",
                Location = new Point(20, 110),
                Size = new Size(450, 25),
                Font = new Font("Microsoft YaHei", 9F),
                ReadOnly = true
            };
            this.Controls.Add(txtOutputFolder);

            var btnSelectOutput = new Button
            {
                Name = "btnSelectOutput",
                Text = "浏览...",
                Location = new Point(480, 109),
                Size = new Size(80, 27),
                Font = new Font("Microsoft YaHei", 9F),
                UseVisualStyleBackColor = true
            };
            btnSelectOutput.Click += BtnSelectOutput_Click;
            this.Controls.Add(btnSelectOutput);

            // 选项
            var chkIncludeSubfolders = new CheckBox
            {
                Name = "chkIncludeSubfolders",
                Text = "包含子文件夹",
                Location = new Point(20, 150),
                Size = new Size(120, 20),
                Font = new Font("Microsoft YaHei", 9F),
                Checked = true
            };
            this.Controls.Add(chkIncludeSubfolders);

            var chkOverwriteExisting = new CheckBox
            {
                Name = "chkOverwriteExisting",
                Text = "覆盖已存在的文件",
                Location = new Point(150, 150),
                Size = new Size(150, 20),
                Font = new Font("Microsoft YaHei", 9F),
                Checked = false
            };
            this.Controls.Add(chkOverwriteExisting);

            var chkDeleteAfterConvert = new CheckBox
            {
                Name = "chkDeleteAfterConvert",
                Text = "删除源XLS",
                Location = new Point(320, 150),
                Size = new Size(100, 20),
                Font = new Font("Microsoft YaHei", 9F),
                Checked = false
            };
            this.Controls.Add(chkDeleteAfterConvert);

            // 转换按钮
            var btnConvert = new Button
            {
                Name = "btnConvert",
                Text = "开始转换",
                Location = new Point(20, 185),
                Size = new Size(100, 35),
                Font = new Font("Microsoft YaHei", 10F, FontStyle.Bold),
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false
            };
            btnConvert.Click += BtnConvert_Click;
            this.Controls.Add(btnConvert);

            var btnCancel = new Button
            {
                Name = "btnCancel",
                Text = "取消",
                Location = new Point(130, 185),
                Size = new Size(80, 35),
                Font = new Font("Microsoft YaHei", 10F),
                UseVisualStyleBackColor = true,
                Enabled = false
            };
            btnCancel.Click += BtnCancel_Click;
            this.Controls.Add(btnCancel);

            var chkAutoMonitor = new CheckBox
            {
                Name = "chkAutoMonitor",
                Text = "自动监控新XLS文件",
                Location = new Point(430, 150),
                Size = new Size(200, 20),
                Font = new Font("Microsoft YaHei", 9F),
                Checked = false
            };
            chkAutoMonitor.CheckedChanged += ChkAutoMonitor_CheckedChanged;
            this.Controls.Add(chkAutoMonitor);

            var lblAutoStatus = new Label
            {
                Name = "lblAutoStatus",
                Text = "自动监控未开启",
                Location = new Point(430, 175),
                Size = new Size(230, 20),
                Font = new Font("Microsoft YaHei", 8.5F),
                ForeColor = Color.Gray
            };
            this.Controls.Add(lblAutoStatus);

            // 进度条
            var progressBar = new ProgressBar
            {
                Name = "progressBar",
                Location = new Point(20, 235),
                Size = new Size(540, 25),
                Style = ProgressBarStyle.Continuous
            };
            this.Controls.Add(progressBar);

            var lblProgress = new Label
            {
                Name = "lblProgress",
                Text = "准备就绪",
                Location = new Point(20, 270),
                Size = new Size(540, 20),
                Font = new Font("Microsoft YaHei", 9F),
                ForeColor = Color.DarkBlue
            };
            this.Controls.Add(lblProgress);

            // 日志文本框
            var txtLog = new TextBox
            {
                Name = "txtLog",
                Location = new Point(20, 300),
                Size = new Size(540, 140),
                Font = new Font("Microsoft YaHei", 8.5F),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                BackColor = Color.White
            };
            this.Controls.Add(txtLog);
        }

        private void BtnSelectInput_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择包含XLS文件的文件夹";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var txtInputFolder = this.Controls["txtInputFolder"] as TextBox;
                    txtInputFolder.Text = dialog.SelectedPath;
                }
            }
        }

        private void BtnSelectOutput_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择输出XLSX文件的文件夹";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    var txtOutputFolder = this.Controls["txtOutputFolder"] as TextBox;
                    txtOutputFolder.Text = dialog.SelectedPath;
                }
            }
        }

        private async void BtnConvert_Click(object sender, EventArgs e)
        {
            var txtInputFolder = this.Controls["txtInputFolder"] as TextBox;
            var txtOutputFolder = this.Controls["txtOutputFolder"] as TextBox;
            var chkIncludeSubfolders = this.Controls["chkIncludeSubfolders"] as CheckBox;
            var chkOverwriteExisting = this.Controls["chkOverwriteExisting"] as CheckBox;
            var btnConvert = this.Controls["btnConvert"] as Button;
            var btnCancel = this.Controls["btnCancel"] as Button;

            if (!TryValidateFolders(txtInputFolder.Text, txtOutputFolder.Text))
            {
                return;
            }

            try
            {
                _isConverting = true;
                btnConvert.Enabled = false;
                btnCancel.Enabled = true;
                ToggleAutoMonitorControls(isEnabled: false);

                var options = new ConversionOptions
                {
                    InputFolder = txtInputFolder.Text,
                    OutputFolder = txtOutputFolder.Text,
                    IncludeSubfolders = chkIncludeSubfolders.Checked,
                    OverwriteExisting = chkOverwriteExisting.Checked,
                    DeleteSourceAfterSuccess = (this.Controls["chkDeleteAfterConvert"] as CheckBox)?.Checked == true
                };

                await ConvertFilesAsync(options);
            }
            catch (Exception ex)
            {
                LogMessage($"转换过程中发生错误: {ex.Message}", Color.Red);
                MessageBox.Show($"转换失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _isConverting = false;
                btnConvert.Enabled = true;
                btnCancel.Enabled = false;
                ToggleAutoMonitorControls(isEnabled: true);
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            _converter.CancelConversion();
            _isConverting = false;
            
            var btnConvert = this.Controls["btnConvert"] as Button;
            var btnCancel = this.Controls["btnCancel"] as Button;
            var lblProgress = this.Controls["lblProgress"] as Label;

            btnConvert.Enabled = true;
            btnCancel.Enabled = false;
            lblProgress.Text = "转换已取消";

            LogMessage("用户取消了转换操作", Color.Orange);
        }

        private bool TryValidateFolders(string input, string output)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                MessageBox.Show("请选择输入文件夹！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (string.IsNullOrWhiteSpace(output))
            {
                MessageBox.Show("请选择输出文件夹！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (!Directory.Exists(input))
            {
                MessageBox.Show("输入文件夹不存在！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (!Directory.Exists(output))
            {
                try
                {
                    Directory.CreateDirectory(output);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"无法创建输出文件夹：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }

            //if (Path.GetFullPath(input).Equals(Path.GetFullPath(output), StringComparison.OrdinalIgnoreCase))
            //{
            //    MessageBox.Show("输入与输出文件夹不能相同！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return false;
            //}

            return true;
        }

        private async Task ConvertFilesAsync(ConversionOptions options)
        {
            var progressBar = this.Controls["progressBar"] as ProgressBar;
            var lblProgress = this.Controls["lblProgress"] as Label;

            var xlsFiles = GetPendingFiles(options);
            if (xlsFiles.Length == 0)
            {
                MessageBox.Show("在指定文件夹中未找到XLS文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            LogMessage($"找到 {xlsFiles.Length} 个XLS文件，开始转换...", Color.Blue);

            progressBar.Maximum = xlsFiles.Length;
            progressBar.Value = 0;

            int successCount = 0;
            int failureCount = 0;

            for (int i = 0; i < xlsFiles.Length && _isConverting; i++)
            {
                var xlsFile = xlsFiles[i];
                var fileName = Path.GetFileNameWithoutExtension(xlsFile);
                var relativePath = Path.GetRelativePath(options.InputFolder, Path.GetDirectoryName(xlsFile));
                var outputDir = Path.Combine(options.OutputFolder, relativePath);
                var xlsxFile = Path.Combine(outputDir, fileName + ".xlsx");

                try
                {
                    lblProgress.Text = $"正在转换: {Path.GetFileName(xlsFile)} ({i + 1}/{xlsFiles.Length})";

                    Directory.CreateDirectory(outputDir);

                    if (File.Exists(xlsxFile) && !options.OverwriteExisting)
                    {
                        LogMessage($"跳过: {Path.GetFileName(xlsxFile)} (文件已存在)", Color.Orange);
                        continue;
                    }

                    await Task.Run(() =>
                    {
                        _converter.EnsureExcelApplication();
                        _converter.ConvertXlsToXlsx(xlsFile, xlsxFile);
                    });

                    if (options.DeleteSourceAfterSuccess)
                    {
                        TryDeleteSourceFile(xlsFile);
                    }

                    successCount++;
                    LogMessage($"成功: {Path.GetFileName(xlsFile)} → {Path.GetFileName(xlsxFile)}", Color.Green);
                }
                catch (Exception ex)
                {
                    failureCount++;
                    LogMessage($"失败: {Path.GetFileName(xlsFile)} - {ex.Message}", Color.Red);
                }

                progressBar.Value = i + 1;
                Application.DoEvents();
            }

            if (_isConverting)
            {
                lblProgress.Text = $"转换完成! 成功: {successCount}, 失败: {failureCount}";
                LogMessage($"批量转换完成! 成功转换 {successCount} 个文件，失败 {failureCount} 个文件", Color.Blue);

                if (successCount > 0)
                {
                    MessageBox.Show($"转换完成!\n成功: {successCount} 个文件\n失败: {failureCount} 个文件", 
                                  "转换完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void LogMessage(string message, Color color)
        {
            var txtLog = this.Controls["txtLog"] as TextBox;
            var timestamp = DateTime.Now.ToString("HH:mm:ss");

            if (txtLog.InvokeRequired)
            {
                txtLog.Invoke(new Action(() =>
                {
                    txtLog.AppendText($"[{timestamp}] {message}\r\n");
                    txtLog.SelectionStart = txtLog.Text.Length;
                    txtLog.ScrollToCaret();
                }));
            }
            else
            {
                txtLog.AppendText($"[{timestamp}] {message}\r\n");
                txtLog.SelectionStart = txtLog.Text.Length;
                txtLog.ScrollToCaret();
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (_isConverting)
            {
                var result = MessageBox.Show("转换正在进行中，确定要退出吗？", "确认退出", 
                                           MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.No)
                {
                    e.Cancel = true;
                    return;
                }
                _converter.CancelConversion();
            }

            StopAutoMonitor();
            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
            }
            _trayMenu?.Dispose();
            _converter?.Dispose();
            _conversionSemaphore.Dispose();
            base.OnFormClosing(e);
        }

        private void InitializeTrayIcon()
        {
            _trayMenu = new ContextMenuStrip();
            _trayMenu.Items.Add("显示主窗口", null, (s, e) => RestoreFromTray());
            _trayMenu.Items.Add("退出", null, (s, e) =>
            {
                _notifyIcon.Visible = false;
                Close();
            });

            _notifyIcon = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Visible = true,
                Text = "XLS到XLSX批量转换器",
                ContextMenuStrip = _trayMenu
            };

            _notifyIcon.DoubleClick += (s, e) => RestoreFromTray();
        }

        private void MainForm_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                Hide();
                _notifyIcon.Visible = true;

                if (!_hasShownTrayTip)
                {
                    _notifyIcon.ShowBalloonTip(
                        3000,
                        "XLS到XLSX批量转换器",
                        "程序最小化到了系统托盘，双击图标可恢复",
                        ToolTipIcon.Info);
                    _hasShownTrayTip = true;
                }
            }
        }

        private void RestoreFromTray()
        {
            Show();
            WindowState = FormWindowState.Normal;
            Activate();
        }

        private void ChkAutoMonitor_CheckedChanged(object sender, EventArgs e)
        {
            if (_suppressAutoMonitorEvent)
            {
                return;
            }

            if (!(sender is CheckBox chkAutoMonitor))
            {
                return;
            }

            if (chkAutoMonitor.Checked)
            {
                StartAutoMonitor();
            }
            else
            {
                StopAutoMonitor(userInitiated: true);
            }
        }

        private void TryDeleteSourceFile(string xlsFile)
        {
            try
            {
                File.Delete(xlsFile);
                LogMessage($"已删除源文件: {Path.GetFileName(xlsFile)}", Color.Gray);
            }
            catch (Exception ex)
            {
                LogMessage($"删除源文件失败: {Path.GetFileName(xlsFile)} - {ex.Message}", Color.Red);
            }
        }

        private void StartAutoMonitor()
        {
            if (_isAutoMonitoring)
            {
                return;
            }

            var txtInputFolder = this.Controls["txtInputFolder"] as TextBox;
            var txtOutputFolder = this.Controls["txtOutputFolder"] as TextBox;
            var chkIncludeSubfolders = this.Controls["chkIncludeSubfolders"] as CheckBox;
            var chkOverwriteExisting = this.Controls["chkOverwriteExisting"] as CheckBox;
            var chkDeleteAfterConvert = this.Controls["chkDeleteAfterConvert"] as CheckBox;
            var chkAutoMonitor = this.Controls["chkAutoMonitor"] as CheckBox;
            var lblAutoStatus = this.Controls["lblAutoStatus"] as Label;

            if (!TryValidateFolders(txtInputFolder.Text, txtOutputFolder.Text))
            {
                _suppressAutoMonitorEvent = true;
                chkAutoMonitor.Checked = false;
                _suppressAutoMonitorEvent = false;
                return;
            }

            _autoModeOptions = new ConversionOptions
            {
                InputFolder = txtInputFolder.Text,
                OutputFolder = txtOutputFolder.Text,
                IncludeSubfolders = chkIncludeSubfolders.Checked,
                OverwriteExisting = chkOverwriteExisting.Checked,
                DeleteSourceAfterSuccess = chkDeleteAfterConvert != null && chkDeleteAfterConvert.Checked
            };

            _watcherCancellation = new CancellationTokenSource();
            _fileWatcher = CreateFileWatcher(_autoModeOptions.InputFolder, _autoModeOptions.IncludeSubfolders);
            _fileWatcher.EnableRaisingEvents = true;

            _backgroundProcessorTask = Task.Run(() => ProcessPendingFilesAsync(_watcherCancellation.Token));

            _isAutoMonitoring = true;
            lblAutoStatus.Text = "自动监控已开启";
            lblAutoStatus.ForeColor = Color.Green;
            LogMessage("自动监控已开启，将自动转换新XLS文件", Color.Blue);
        }

        private void StopAutoMonitor(bool userInitiated = false)
        {
            if (!_isAutoMonitoring)
            {
                return;
            }

            _isAutoMonitoring = false;

            try
            {
                _watcherCancellation?.Cancel();
                _fileWatcher?.Dispose();
                _fileWatcher = null;
            }
            catch (Exception ex)
            {
                LogMessage($"停止自动监控时发生错误: {ex.Message}", Color.Red);
            }

            if (_backgroundProcessorTask != null)
            {
                try
                {
                    _backgroundProcessorTask.Wait(TimeSpan.FromSeconds(2));
                }
                catch (AggregateException)
                {
                    // 忽略后台任务关闭时的异常
                }
            }

            _backgroundProcessorTask = null;

            _watcherCancellation?.Dispose();
            _watcherCancellation = null;

            _autoRetryCounts.Clear();
            _queuedFiles.Clear();
            while (_pendingFiles.TryDequeue(out _)) { }

            var chkAutoMonitor = this.Controls["chkAutoMonitor"] as CheckBox;
            var lblAutoStatus = this.Controls["lblAutoStatus"] as Label;

            if (chkAutoMonitor != null)
            {
                _suppressAutoMonitorEvent = true;
                try
                {
                    chkAutoMonitor.Checked = false;
                }
                finally
                {
                    _suppressAutoMonitorEvent = false;
                }
            }

            if (lblAutoStatus != null)
            {
                lblAutoStatus.Text = userInitiated ? "自动监控已关闭" : "自动监控未开启";
                lblAutoStatus.ForeColor = Color.Gray;
            }

            LogMessage("自动监控已关闭", Color.Blue);
        }

        private FileSystemWatcher CreateFileWatcher(string folder, bool includeSubfolders)
        {
            var watcher = new FileSystemWatcher(folder)
            {
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite,
                Filter = "*.xls",
                IncludeSubdirectories = includeSubfolders,
                EnableRaisingEvents = false
            };

            watcher.Created += (s, e) => EnqueueFileForProcessing(e.FullPath);
            watcher.Changed += (s, e) => EnqueueFileForProcessing(e.FullPath);
            watcher.Renamed += (s, e) => EnqueueFileForProcessing(e.FullPath);

            return watcher;
        }

        private void EnqueueFileForProcessing(string filePath)
        {
            if (!_isAutoMonitoring)
            {
                return;
            }

            if (!filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                Path.GetFileName(filePath).StartsWith("~$"))
            {
                return;
            }

            var normalizedPath = Path.GetFullPath(filePath);
            _queuedFiles[normalizedPath] = 0;
            _pendingFiles.Enqueue(normalizedPath);
        }

        private async Task ProcessPendingFilesAsync(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                if (_pendingFiles.TryDequeue(out var filePath))
                {
                    if (!ShouldProcessFile(filePath))
                    {
                        continue;
                    }

                    try
                    {
                        await _conversionSemaphore.WaitAsync(cancellationToken);
                        await ConvertSingleFileAsync(filePath, _autoModeOptions, cancellationToken);
                    }
                    catch (OperationCanceledException)
                    {
                        return;
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"自动转换失败: {Path.GetFileName(filePath)} - {ex.Message}", Color.Red);
                        ScheduleRetry(filePath);
                    }
                    finally
                    {
                        _conversionSemaphore.Release();
                    }
                }
                else
                {
                    await Task.Delay(500, cancellationToken);
                }
            }
        }

        private bool ShouldProcessFile(string filePath)
        {
            if (!_queuedFiles.TryGetValue(filePath, out _))
            {
                return false;
            }

            if (!File.Exists(filePath))
            {
                _queuedFiles.TryRemove(filePath, out _);
                return false;
            }

            var lastWriteTime = File.GetLastWriteTimeUtc(filePath);
            if (DateTime.UtcNow - lastWriteTime < TimeSpan.FromSeconds(2))
            {
                _queuedFiles[filePath] = 0;
                _pendingFiles.Enqueue(filePath);
                return false;
            }

            return true;
        }

        private async Task ConvertSingleFileAsync(string xlsFile, ConversionOptions options, CancellationToken cancellationToken)
        {
            var relativePath = Path.GetRelativePath(options.InputFolder, Path.GetDirectoryName(xlsFile));
            var outputDir = Path.Combine(options.OutputFolder, relativePath);
            var fileName = Path.GetFileNameWithoutExtension(xlsFile);
            var xlsxFile = Path.Combine(outputDir, fileName + ".xlsx");

            try
            {
                Directory.CreateDirectory(outputDir);

                if (File.Exists(xlsxFile) && !options.OverwriteExisting)
                {
                    LogMessage($"自动转换跳过: {Path.GetFileName(xlsxFile)} (文件已存在)", Color.Orange);
                    _queuedFiles.TryRemove(xlsFile, out _);
                    return;
                }

                LogMessage($"自动转换开始: {Path.GetFileName(xlsFile)}", Color.Blue);
                await Task.Run(() =>
                {
                    _converter.EnsureExcelApplication();
                    _converter.ConvertXlsToXlsx(xlsFile, xlsxFile);
                }, cancellationToken);

                if (options.DeleteSourceAfterSuccess)
                {
                    TryDeleteSourceFile(xlsFile);
                }

                LogMessage($"自动转换成功: {Path.GetFileName(xlsFile)}", Color.Green);
                _queuedFiles.TryRemove(xlsFile, out _);
                _autoRetryCounts.TryRemove(xlsFile, out _);
            }
            catch (Exception ex)
            {
                LogMessage($"自动转换失败: {Path.GetFileName(xlsFile)} - {ex.Message}", Color.Red);
                ScheduleRetry(xlsFile);
            }
        }

        private void ScheduleRetry(string filePath)
        {
            var retryCount = _autoRetryCounts.AddOrUpdate(filePath, 1, (key, current) => current + 1);

            if (retryCount > 3)
            {
                LogMessage($"自动转换超过最大重试次数，已跳过: {Path.GetFileName(filePath)}", Color.Red);
                _queuedFiles.TryRemove(filePath, out _);
                _autoRetryCounts.TryRemove(filePath, out _);
                return;
            }

            Task.Delay(TimeSpan.FromSeconds(Math.Min(30, retryCount * 5))).ContinueWith(_ =>
            {
                if (_isAutoMonitoring)
                {
                    _queuedFiles[filePath] = 0;
                    _pendingFiles.Enqueue(filePath);
                }
            }, TaskScheduler.Default);
        }

        private string[] GetPendingFiles(ConversionOptions options)
        {
            var searchOption = options.IncludeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var discoveredFiles = Directory.GetFiles(options.InputFolder, "*.xls", searchOption)
                                           .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                                           .Select(Path.GetFullPath)
                                           .ToArray();

            if (_isAutoMonitoring)
            {
                var allFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var file in discoveredFiles)
                {
                    allFiles.Add(file);
                }

                while (_pendingFiles.TryDequeue(out var queuedFile))
                {
                    allFiles.Add(Path.GetFullPath(queuedFile));
                }

                return allFiles.ToArray();
            }

            return discoveredFiles;
        }

        private void ToggleAutoMonitorControls(bool isEnabled)
        {
            var chkAutoMonitor = this.Controls["chkAutoMonitor"] as CheckBox;
            var chkIncludeSubfolders = this.Controls["chkIncludeSubfolders"] as CheckBox;
            var chkDeleteAfterConvert = this.Controls["chkDeleteAfterConvert"] as CheckBox;

            if (chkAutoMonitor != null)
            {
                chkAutoMonitor.Enabled = isEnabled;
            }

            if (chkIncludeSubfolders != null)
            {
                chkIncludeSubfolders.Enabled = isEnabled;
            }

            if (chkDeleteAfterConvert != null)
            {
                chkDeleteAfterConvert.Enabled = isEnabled;
            }
        }
    }
}