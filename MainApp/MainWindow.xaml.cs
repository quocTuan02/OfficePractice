using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using CommonLibrary;
using CommonLibrary.Models;
using ScoringEngine;

namespace MainApp
{
    public partial class MainWindow : Window
    {
        #region Win32 API
        [DllImport("user32.dll")] static extern IntPtr SetParent(IntPtr child, IntPtr parent);
        [DllImport("user32.dll")] static extern bool MoveWindow(IntPtr hwnd, int x, int y, int w, int h, bool repaint);
        [DllImport("user32.dll")] static extern int SetWindowLong(IntPtr hwnd, int idx, int val);
        [DllImport("user32.dll")] static extern int GetWindowLong(IntPtr hwnd, int idx);
        [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int w, int h, uint flags);
        [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hwnd, int cmd);
        [DllImport("user32.dll")] static extern bool ScreenToClient(IntPtr hwnd, ref POINT pt);
        [DllImport("user32.dll")] static extern bool InvalidateRect(IntPtr hwnd, IntPtr rect, bool erase);
        [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
        [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr hwnd, out RECT rect);
        [DllImport("user32.dll")] static extern bool SetWindowRgn(IntPtr hwnd, IntPtr hRgn, bool redraw);
        [DllImport("user32.dll")] static extern int GetSystemMetrics(int nIndex);
        [DllImport("gdi32.dll")]  static extern IntPtr CreateRectRgn(int x1, int y1, int x2, int y2);

        [StructLayout(LayoutKind.Sequential)]
        struct RECT { public int Left, Top, Right, Bottom; }

        const int SM_CYCAPTION  = 4;
        const int SM_CYSIZEFRAME = 33;

        [StructLayout(LayoutKind.Sequential)]
        struct POINT { public int x, y; }

        const int GWL_STYLE       = -16;
        const int WS_CAPTION      = 0x00C00000;
        const int WS_THICKFRAME   = 0x00040000;
        const int WS_SYSMENU      = 0x00080000;
        const int WS_CLIPSIBLINGS = 0x04000000;
        const int WS_CHILD        = 0x40000000;
        const int WS_POPUP        = unchecked((int)0x80000000);
        const int SW_SHOW         = 5;
        const uint SWP_FRAMECHANGED = 0x0020;
        const uint SWP_NOSIZE     = 0x0001;
        const uint SWP_NOMOVE     = 0x0002;
        const uint SWP_SHOWWINDOW  = 0x0040;
        const uint SWP_NOZORDER    = 0x0004;
        static readonly IntPtr HWND_TOP       = IntPtr.Zero;
        static readonly IntPtr HWND_TOPMOST   = new IntPtr(-1);
        static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
        #endregion

        private PipeServer _pipeServer;
        private IntPtr _officeHwnd = IntPtr.Zero;
        private Process _officeProcess;
        private bool _isExcel;

        // Package integration
        private TestModel _loadedTest;
        private TaskModel _currentTask;
        private string _packageDir = string.Empty;
        private readonly ScoringExecutor _executor = new ScoringExecutor();

        public MainWindow()
        {
            InitializeComponent();
            StartPipeServer();
            // Khi MainApp được kích hoạt lại, đẩy Excel lên trên (top-level window)
            Activated += (s, e) =>
            {
                if (_officeHwnd != IntPtr.Zero && _isExcel)
                    SetWindowPos(_officeHwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE | SWP_SHOWWINDOW);
            };
        }

        private void StartPipeServer()
        {
            _pipeServer = new PipeServer();
            _pipeServer.MessageReceived += msg =>
                Dispatcher.Invoke(() => AppendLog($"[{msg.Application}] {msg.EventType}: {msg.Detail}"));
            _pipeServer.Start();
        }

        private async void OpenExcel_Click(object sender, RoutedEventArgs e)
        {
            _isExcel = true;
            await OpenOffice("Excel.Application", "EXCEL", "Excel — Nhập liệu cơ bản");
        }

        private async void OpenWord_Click(object sender, RoutedEventArgs e)
        {
            _isExcel = false;
            await OpenOffice("Word.Application", "WINWORD", "Word — Soạn thảo văn bản");
        }

        private async Task OpenOffice(string progId, string processName, string title)
        {
            BtnExcel.IsEnabled = false;
            BtnWord.IsEnabled = false;
            AppendLog($"Đang mở {processName}...");

            try
            {
                var type = Type.GetTypeFromProgID(progId, throwOnError: true);
                dynamic app = Activator.CreateInstance(type);
                app.Visible = true;
                if (progId.Contains("Excel")) app.Workbooks.Add();
                else app.Documents.Add();

                await Task.Delay(2000);

                var procs = Process.GetProcessesByName(processName);
                if (procs.Length == 0) throw new Exception($"Không tìm thấy process {processName}");
                _officeProcess = procs[procs.Length - 1];

                for (int i = 0; i < 10; i++)
                {
                    _officeHwnd = _officeProcess.MainWindowHandle;
                    if (_officeHwnd != IntPtr.Zero) break;
                    await Task.Delay(400);
                    _officeProcess.Refresh();
                }

                if (_officeHwnd == IntPtr.Zero) throw new Exception("Không lấy được window handle.");

                EmbedOfficeWindow();
                ExerciseTitle.Text = title;
                OfficePlaceholder.Visibility = Visibility.Collapsed;
                BtnClose.IsEnabled = true;
                AppendLog("Đã nhúng vào cửa sổ thành công.");
            }
            catch (Exception ex)
            {
                AppendLog($"Lỗi: {ex.Message}");
                BtnExcel.IsEnabled = true;
                BtnWord.IsEnabled = true;
            }
        }

        private void EmbedOfficeWindow()
        {
            if (_officeHwnd == IntPtr.Zero) return;

            if (_isExcel)
            {
                // Excel: giữ top-level (DWM/ribbon không hoạt động với WS_CHILD)
                // Chỉ xóa border + sysmenu, GIỮ WS_CAPTION để ribbon render được
                int style = GetWindowLong(_officeHwnd, GWL_STYLE);
                style &= ~(WS_THICKFRAME | WS_SYSMENU | WS_POPUP);
                SetWindowLong(_officeHwnd, GWL_STYLE, style);
                SetWindowPos(_officeHwnd, HWND_TOP, 0, 0, 0, 0,
                    SWP_FRAMECHANGED | SWP_NOSIZE | SWP_NOMOVE);
            }
            else
            {
                // Word (SDI): SetParent hoạt động tốt
                int style = GetWindowLong(_officeHwnd, GWL_STYLE);
                style &= ~(WS_CAPTION | WS_THICKFRAME | WS_SYSMENU | WS_POPUP);
                style |= WS_CHILD | WS_CLIPSIBLINGS;
                SetWindowLong(_officeHwnd, GWL_STYLE, style);
                SetWindowPos(_officeHwnd, HWND_TOP, 0, 0, 0, 0,
                    SWP_FRAMECHANGED | SWP_NOSIZE | SWP_NOMOVE);
                var mainHwnd = new WindowInteropHelper(this).Handle;
                SetParent(_officeHwnd, mainHwnd);
            }

            PositionOfficeWindow();
            ShowWindow(_officeHwnd, SW_SHOW);
            InvalidateRect(_officeHwnd, IntPtr.Zero, true);
        }

        private void PositionOfficeWindow()
        {
            if (_officeHwnd == IntPtr.Zero) return;

            Dispatcher.Invoke(() =>
            {
                try
                {
                    var dpi = VisualTreeHelper.GetDpi(OfficePanelBorder);
                    var screenPos = OfficePanelBorder.PointToScreen(new Point(0, 0));
                    var w = (int)(OfficePanelBorder.ActualWidth * dpi.DpiScaleX);
                    var h = (int)(OfficePanelBorder.ActualHeight * dpi.DpiScaleY);

                    if (_isExcel)
                    {
                        // Tính chiều cao title bar (có DPI scale)
                        var captionH = (int)(GetSystemMetrics(SM_CYCAPTION) * dpi.DpiScaleY);

                        // Dịch Excel lên trên captionH px để title bar ẩn sau toolbar của app
                        MoveWindow(_officeHwnd, (int)screenPos.X, (int)screenPos.Y - captionH,
                            w, h + captionH, true);

                        // Clip: ẩn phần title bar (y < captionH), chỉ hiện từ captionH xuống
                        var rgn = CreateRectRgn(0, captionH, w, h + captionH);
                        SetWindowRgn(_officeHwnd, rgn, true);

                        // Đưa Excel lên trên cùng (trên MainApp)
                        SetWindowPos(_officeHwnd, HWND_TOP, 0, 0, 0, 0,
                            SWP_NOSIZE | SWP_NOMOVE | SWP_SHOWWINDOW);
                    }
                    else
                    {
                        // Word: dùng client coords (true child window)
                        var mainHwnd = new WindowInteropHelper(this).Handle;
                        var pt = new POINT { x = (int)screenPos.X, y = (int)screenPos.Y };
                        ScreenToClient(mainHwnd, ref pt);
                        MoveWindow(_officeHwnd, pt.x, pt.y, w, h, true);
                    }
                }
                catch { }
            });
        }

        private void OnWindowMoved(object sender, EventArgs e) => PositionOfficeWindow();
        private void OnWindowResized(object sender, SizeChangedEventArgs e) => PositionOfficeWindow();

        private void CloseOffice_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_officeHwnd != IntPtr.Zero)
                {
                    // Xóa region clip trước khi kill để tránh artifact
                    SetWindowRgn(_officeHwnd, IntPtr.Zero, false);
                    if (!_isExcel) SetParent(_officeHwnd, IntPtr.Zero);
                    _officeHwnd = IntPtr.Zero;
                }
                _officeProcess?.Kill();
                _officeProcess = null;
                OfficePlaceholder.Visibility = Visibility.Visible;
                BtnExcel.IsEnabled = true;
                BtnWord.IsEnabled = true;
                BtnClose.IsEnabled = false;
                AppendLog("Office đã đóng.");
            }
            catch { }
        }

        private void ClearLog_Click(object sender, RoutedEventArgs e) => LogText.Text = string.Empty;

        // ── Package integration ────────────────────────────────────────

        private void LoadPackage_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Package ZIP|*.zip|All Files|*.*",
                Title  = "Chọn package bài thi (.zip)"
            };
            if (dlg.ShowDialog() != true) return;

            try
            {
                // Clean up previous extraction
                if (!string.IsNullOrEmpty(_packageDir) && Directory.Exists(_packageDir))
                    Directory.Delete(_packageDir, recursive: true);

                _packageDir = Path.Combine(Path.GetTempPath(),
                    "MOS_Pkg_" + Guid.NewGuid().ToString("N")[..8]);
                ZipFile.ExtractToDirectory(dlg.FileName, _packageDir);

                var testJson = Directory
                    .GetFiles(_packageDir, "test.json", SearchOption.AllDirectories)
                    .FirstOrDefault();

                if (testJson == null)
                    throw new FileNotFoundException("Không tìm thấy test.json trong package.");

                _loadedTest = JsonSerializer.Deserialize<TestModel>(File.ReadAllText(testJson));
                if (_loadedTest == null) throw new InvalidDataException("test.json không hợp lệ.");

                // Resolve template file paths relative to package Templates/ folder
                var templatesDir = Path.Combine(Path.GetDirectoryName(testJson)!, "Templates");
                foreach (var task in _loadedTest.Tasks)
                {
                    if (!string.IsNullOrEmpty(task.TemplateFile))
                    {
                        var fileName = Path.GetFileName(task.TemplateFile);
                        var localPath = Path.Combine(templatesDir, fileName);
                        if (File.Exists(localPath))
                            task.TemplateFile = localPath;
                    }
                }

                // Populate task selector
                TaskComboBox.Items.Clear();
                foreach (var task in _loadedTest.Tasks)
                    TaskComboBox.Items.Add(new ComboBoxItem
                    {
                        Content = $"Task {task.Number}: {task.Objective}",
                        Tag     = task
                    });

                TaskComboBox.SelectedIndex = 0;
                TaskSelectorBorder.Visibility = Visibility.Visible;
                AppendLog($"Đã load: {_loadedTest.Name} · {_loadedTest.Tasks.Count} tasks · {_loadedTest.TotalPoints} điểm");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi load package:\n{ex.Message}", "Lỗi",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void TaskComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (TaskComboBox.SelectedItem is ComboBoxItem item && item.Tag is TaskModel task)
            {
                _currentTask = task;
                UpdateTaskView();
            }
        }

        private void UpdateTaskView()
        {
            if (_currentTask == null) return;

            ExerciseTitle.Text = $"Task {_currentTask.Number}: {_currentTask.Objective}";
            TaskPointsLabel.Text = $"{_currentTask.Points} điểm · {_currentTask.ScoringRules.Count} rules";

            // Build dynamic instruction steps
            StepsPanel.Children.Clear();
            var lines = _currentTask.Instruction
                .Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < lines.Length; i++)
            {
                var row = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Margin = new Thickness(0, 0, 0, 6)
                };

                var badge = new Border
                {
                    Width = 20, Height = 20,
                    CornerRadius = new CornerRadius(10),
                    Background = new SolidColorBrush(Color.FromRgb(0x31, 0x32, 0x44)),
                    Margin = new Thickness(0, 2, 8, 0),
                    Child = new TextBlock
                    {
                        Text = (i + 1).ToString(),
                        Foreground = new SolidColorBrush(Color.FromRgb(0xCD, 0xD6, 0xF4)),
                        FontSize = 11,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment   = VerticalAlignment.Center
                    }
                };

                var text = new TextBlock
                {
                    Text = lines[i].Trim(),
                    Foreground = new SolidColorBrush(Color.FromRgb(0xBA, 0xC2, 0xDE)),
                    FontSize = 12,
                    VerticalAlignment = VerticalAlignment.Top,
                    TextWrapping = TextWrapping.Wrap,
                    MaxWidth = 280
                };

                row.Children.Add(badge);
                row.Children.Add(text);
                StepsPanel.Children.Add(row);
            }

            // Auto-open template file if it exists and Excel is not already open
            if (!string.IsNullOrEmpty(_currentTask.TemplateFile) &&
                File.Exists(_currentTask.TemplateFile) &&
                _officeHwnd == IntPtr.Zero)
            {
                AppendLog($"Template: {Path.GetFileName(_currentTask.TemplateFile)}");
                BtnSubmit.Content = "Lưu & Nộp bài";
            }
        }

        private async void SubmitAndScore_Click(object sender, RoutedEventArgs e)
        {
            if (_currentTask == null || _currentTask.ScoringRules.Count == 0)
            {
                AppendLog("Đã nộp bài (không có scoring rules).");
                MessageBox.Show("Đã nộp bài!", "Thông báo",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Ask for the answer file
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xlsm;*.xls|All Files|*.*",
                Title  = "Chọn file bài làm của bạn để chấm điểm"
            };
            if (dlg.ShowDialog() != true) return;

            BtnSubmit.IsEnabled = false;
            BtnSubmit.Content   = "Đang chấm...";
            AppendLog($"Chấm điểm: {Path.GetFileName(dlg.FileName)}");

            try
            {
                var result = await Task.Run(
                    () => _executor.ExecuteTask(_currentTask, dlg.FileName));

                AppendLog($"Kết quả Task {result.TaskNumber}: {result.TotalPoints}/{result.MaxPoints} điểm");

                var win = new ScoringResultWindow(result) { Owner = this };
                win.ShowDialog();
            }
            catch (Exception ex)
            {
                AppendLog($"Lỗi chấm: {ex.Message}");
                MessageBox.Show($"Không thể chấm điểm:\n{ex.Message}", "Lỗi",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnSubmit.IsEnabled = true;
                BtnSubmit.Content   = "Lưu & Nộp bài";
            }
        }

        private void AppendLog(string message)
        {
            LogText.Text += $"[{DateTime.Now:HH:mm:ss}] {message}\n";
            LogScrollViewer?.ScrollToEnd();
        }

        protected override void OnClosed(EventArgs e)
        {
            _pipeServer?.Stop();
            _pipeServer?.Dispose();
            base.OnClosed(e);
        }
    }
}
