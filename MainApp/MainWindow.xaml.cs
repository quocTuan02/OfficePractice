using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using CommonLibrary;

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
        static readonly IntPtr HWND_TOPMOST   = new IntPtr(-1);
        static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
        #endregion

        private PipeServer _pipeServer;
        private IntPtr _officeHwnd = IntPtr.Zero;
        private Process _officeProcess;
        private bool _isExcel;

        public MainWindow()
        {
            InitializeComponent();
            StartPipeServer();
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

                // Với Excel, tìm XLDESK window để nhúng chính xác hơn
                if (_isExcel)
                {
                    // Dùng main window cho Excel, không thêm WS_CHILD vì MDI không tương thích
                }

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

            int style = GetWindowLong(_officeHwnd, GWL_STYLE);

            if (_isExcel)
            {
                // Excel: KHÔNG dùng SetParent — SetParent làm ribbon bị ẩn
                // Chỉ xóa decoration rồi snap vào vị trí panel (top-level window)
                int exStyle = GetWindowLong(_officeHwnd, GWL_STYLE);
                exStyle &= ~(WS_THICKFRAME | WS_SYSMENU | WS_POPUP);
                SetWindowLong(_officeHwnd, GWL_STYLE, exStyle);
                SetWindowPos(_officeHwnd, IntPtr.Zero, 0, 0, 0, 0,
                    SWP_FRAMECHANGED | SWP_NOSIZE | SWP_NOMOVE);
            }
            else
            {
                // Word (SDI): WS_CHILD hoạt động tốt
                style &= ~(WS_CAPTION | WS_THICKFRAME | WS_SYSMENU | WS_POPUP);
                style |= WS_CHILD | WS_CLIPSIBLINGS;
                SetWindowLong(_officeHwnd, GWL_STYLE, style);
                SetWindowPos(_officeHwnd, IntPtr.Zero, 0, 0, 0, 0,
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
                        // Excel là top-level window — đặt tại vị trí panel, Excel ở trên MainApp
                        MoveWindow(_officeHwnd, (int)screenPos.X, (int)screenPos.Y, w, h, true);
                        // Đặt Excel TRÊN MainApp để hiển thị (không bị che)
                        var mainHwnd2 = new WindowInteropHelper(this).Handle;
                        SetWindowPos(_officeHwnd, mainHwnd2, 0, 0, 0, 0,
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
                    SetParent(_officeHwnd, IntPtr.Zero);
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
