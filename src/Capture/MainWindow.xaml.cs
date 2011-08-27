using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using System.Runtime.InteropServices;
using System.Windows.Interop;
using System.Windows.Threading;
//using System.Drawing;

namespace Capture
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        UserSetting setting = UserSetting.LoadSetting();
        System.Windows.Forms.NotifyIcon _notifyIcon;
        bool _exitClick = false;

        System.Collections.ObjectModel.ObservableCollection<CaptureEntry> _captureList;
        SettingWindow w;
        MouseHook m;
        System.Windows.Automation.TreeWalker tw = System.Windows.Automation.TreeWalker.ControlViewWalker;

        public MainWindow()
        {
            InitializeComponent();

            //System.Windows.Point point = new Point(100, 100);
            //System.Windows.Automation.AutomationElement el = null;
            //el = System.Windows.Automation.AutomationElement.FromPoint(point);
            //System.Windows.Automation.AutomationElement el_top = TopElement(el);

            //Cpt(new Point(200, 300));

            try
            {
                this.ShowInTaskbar = false;
                _notifyIcon = new System.Windows.Forms.NotifyIcon();
                _notifyIcon.Text = "Capture";
                _notifyIcon.Icon = new System.Drawing.Icon("icon.ico");
                _notifyIcon.Visible = true;

                System.Windows.Forms.ContextMenuStrip menuStrip = new System.Windows.Forms.ContextMenuStrip();
                System.Windows.Forms.ToolStripMenuItem openItem = new System.Windows.Forms.ToolStripMenuItem();
                openItem.Text = "Capture を開く";
                openItem.Click += new EventHandler(openItem_Click);
                menuStrip.Items.Add(openItem);

                System.Windows.Forms.ToolStripMenuItem exitItem = new System.Windows.Forms.ToolStripMenuItem();
                exitItem.Text = "終了";
                exitItem.Click += new EventHandler(exitItem_Click);
                menuStrip.Items.Add(exitItem);
                       

                _notifyIcon.ContextMenuStrip = menuStrip;
                _notifyIcon.DoubleClick += new EventHandler(_notifyIcon_DoubleClick);

                this.Closing += new System.ComponentModel.CancelEventHandler(MainWindow_Closing);
                this.Closed += new EventHandler(MainWindow_Closed);

                this.btnStop.IsEnabled = false;


                // WindowのHandleを取得
                WindowInteropHelper _host = new WindowInteropHelper(this);
                this._windowHandle = _host.Handle;

                // HotKeyを設定
                SetupHotKey();
                ComponentDispatcher.ThreadPreprocessMessage
                    += ComponentDispatcher_ThreadPreprocessMessage;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }


        void openItem_Click(object sender, EventArgs e)
        {
            this.Show();
            this.Activate();
        }

        void _notifyIcon_DoubleClick(object sender, EventArgs e)
        {
            openItem_Click(null, null);
        }

        void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (!_exitClick)
            {
                //タスクトレイから終了していない場合は終了しない
                this.Visibility = System.Windows.Visibility.Collapsed;
                e.Cancel = true;
            }
            else
            {
                btnStop_Click(null, null); //記録停止
            }
        }

        void exitItem_Click(object sender, EventArgs e)
        {
            if (w != null)
                w.Close();
            _exitClick = true;
            this.Close();
        }

        void MainWindow_Closed(object sender, EventArgs e)
        {
            // HotKeyの登録削除
            UnregisterHotKey(this._windowHandle, HOTKEY_ID1);
            _notifyIcon.Visible = false;
        }




        private const int HOTKEY_ID1 = 0x0001;
        private const int HOTKEY_ID2 = 0x0002;
        // HotKey Message ID
        private const int WM_HOTKEY = 0x0312;
        private IntPtr _windowHandle;
        [DllImport("user32.dll")]
        private static extern int RegisterHotKey(IntPtr hWnd, int id, int MOD_KEY, int VK);
        [DllImport("user32.dll")]
        private static extern int UnregisterHotKey(IntPtr hWnd, int id); 

        // HotKeyの登録
        private void SetupHotKey()
        {
            RegisterHotKey(this._windowHandle, HOTKEY_ID1, 0,
               KeyInterop.VirtualKeyFromKey(Key.PrintScreen));
        }

        // HotKeyの動作を設定する
        public void ComponentDispatcher_ThreadPreprocessMessage(
        ref MSG msg, ref bool handled)
        {
            if (msg.message == WM_HOTKEY)
            {
                if (msg.wParam.ToInt32() == HOTKEY_ID1)
                {
                    TimeSpan ts = DateTime.Now - lastCaptured;
                    if (ts.TotalSeconds < 0.5)
                        return; //とりあえず0.5秒以内のクリックは無視しておく

                    lastCaptured = DateTime.Now;
                    System.Threading.Thread t = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(DoCapture));
                    t.Start(null);
                }
            }
        }



        /// <summary>
        /// 指定したウィンドウハンドルの親を探す。(同じプロセスIDで一番上のAutomationElement)
        /// </summary>
        /// <param name="el"></param>
        /// <returns></returns>
        System.Windows.Automation.AutomationElement TopElement(System.Windows.Automation.AutomationElement el)
        {
            if (el == null)
                return null;

            System.Windows.Automation.AutomationElement p = tw.GetParent(el);
            if (p == null)
                return el; //親は無い

            if (p.Current.ProcessId == el.Current.ProcessId)
                return TopElement(p);

            return el;
        }

        DateTime lastCaptured = DateTime.Now;
        void m_MouseHooked(object sender, MouseHookedEventArgs e)
        {
            if (e.Message == MouseMessage.LDown || e.Message == MouseMessage.RDown)
            {
                TimeSpan ts = DateTime.Now - lastCaptured;
                if (ts.TotalSeconds < 0.5)
                    return; //とりあえず0.5秒以内のクリックは無視しておく

                lastCaptured = DateTime.Now;
                System.Threading.Thread t = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(DoCapture));
                t.Start(e.Message);
            }
        }

        void DoCapture(object param)
        {
            MouseMessage mouse = (MouseMessage)param;
            if (_captureList.Count > 100)
                return; //ファイル数制限をつけておく

            System.Windows.Point point = new Point(System.Windows.Forms.Cursor.Position.X, System.Windows.Forms.Cursor.Position.Y);
            System.Windows.Automation.AutomationElement el = null;

            try
            {
                el = System.Windows.Automation.AutomationElement.FromPoint(point);

                if (el == null)
                    return;

                if (el.Current.ProcessId == System.Diagnostics.Process.GetCurrentProcess().Id)
                    return; //自分自身のWindow上のクリックは無視

                string path_w = System.IO.Path.Combine(setting.SaveDir, DateTime.Now.ToString("HHmmssfff") + "-w.png");
                string path_c = System.IO.Path.Combine(setting.SaveDir, DateTime.Now.ToString("HHmmssfff") + "-c.png");

                System.Windows.Automation.AutomationElement el_top = TopElement(el);

                CaptureEntry entry = new CaptureEntry();
                entry.MouseMessage = mouse;
                entry.ClickedControlType = el.Current.LocalizedControlType;
                entry.ClickedName = el.Current.Name;
                entry.ClickedRect = el.Current.BoundingRectangle;
                entry.ClickedWindowHandle = el.Current.NativeWindowHandle;
                entry.Cursor = point;
                entry.PathToWindowImage = path_w;
                entry.PathToClipImage = path_c;

                if (el_top != null)
                {
                    entry.WindowControlType = el_top.Current.LocalizedControlType;
                    entry.WindowName = el_top.Current.Name;
                    entry.WindowRect = el_top.Current.BoundingRectangle;
                    entry.WindowHandle = el_top.Current.NativeWindowHandle;
                }

                //キャプチャ
                Cpt(el_top, path_w);
                Cpt(point, path_c, entry);

                this.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate
                {
                    entry.No = _captureList.Count + 1;
                    _captureList.Add(entry);
                }));
            }
            catch (System.Windows.Automation.ElementNotAvailableException)
            {
                //無視
            }
            catch (Exception ex)
            {
                this.Dispatcher.Invoke(DispatcherPriority.Normal, new Action(delegate
                {
                    MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
                }));
            }
        }

        /// <summary>
        /// キャプチャ(ウィンドウ)
        /// </summary>
        /// <param name="el"></param>
        /// <param name="path"></param>
        void Cpt(System.Windows.Automation.AutomationElement el, string path)
        {
            if (el == null)
                return;

            if (el.Current.NativeWindowHandle != 0)
            {
                System.Drawing.Bitmap bmp = CaptureHelper.CaptureWindow((IntPtr)el.Current.NativeWindowHandle);
                bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
            }

            string name = el.Current.ControlType.LocalizedControlType;
            string type = el.Current.ItemType;
            string hnd = el.Current.NativeWindowHandle.ToString();
            string pid = el.Current.ProcessId.ToString();
        }

        /// <summary>
        /// キャプチャ(切り抜き)
        /// </summary>
        /// <param name="p"></param>
        /// <param name="path"></param>
        void Cpt(Point p, string path, CaptureEntry entry)
        {
            int top = setting.ClipSizeTop;
            int left = setting.ClipSizeLeft;
            int right = setting.ClipSizeRight;
            int bottom = setting.ClipSizeBottom;

            System.Drawing.Bitmap bmp = CaptureHelper.CaptureScreen();

            int top_offset = 0;
            int left_offset = 0;

            if ((p.Y + top) < 0)
                top_offset = -((int)(p.Y + top));
            else if (bmp.Height < (p.Y + bottom))
                top_offset = bmp.Height - (int)(p.Y + bottom);

            if ((p.X + left) < 0)
                left_offset = -((int)(p.X + left));
            else if (bmp.Width < (p.X + right))
                left_offset = bmp.Width - (int)(p.X + right);

            int adj_top = top + top_offset;
            int adj_left = left + left_offset;
            int adj_right = right + left_offset;
            int adj_bottom = bottom + top_offset;
            
            //切り抜く
            System.Drawing.Rectangle rect = new System.Drawing.Rectangle((int)(p.X + adj_left), (int)(p.Y + adj_top), -adj_left + adj_right, -adj_top + adj_bottom);

            entry.ClipImageRect = new Rect((int)(p.X + adj_left), (int)(p.Y + adj_top), -adj_left + adj_right, -adj_top + adj_bottom);

            System.Drawing.Bitmap bmpNew = bmp.Clone(rect, bmp.PixelFormat);
            bmpNew.Save(path, System.Drawing.Imaging.ImageFormat.Png);
        }
               

        private void settingMenu_Click(object sender, RoutedEventArgs e)
        {
            if (w != null)
                return;
            w = new SettingWindow();
            w.ShowDialog();
            if (w.DialogResult == true && _captureList != null)
            {
                _captureList.Clear(); //設定を変えたので初期化
            }
            w = null;
        }

        /// <summary>
        /// 記録開始ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!System.IO.Directory.Exists(setting.SaveDir))
                    System.IO.Directory.CreateDirectory(setting.SaveDir);
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存フォルダの作成に失敗しました。\n" + ex.Message);
                return;
            }

            try
            {
                if (_captureList == null)
                {
                    _captureList = new System.Collections.ObjectModel.ObservableCollection<CaptureEntry>();
                    listView1.DataContext = _captureList;
                }

                if (m == null)
                {
                    m = new MouseHook();
                    m.MouseHooked += new MouseHookedEventHandler(m_MouseHooked);
                }
                this.btnStart.IsEnabled = false;
                this.btnStop.IsEnabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }

        /// <summary>
        /// 記録停止ボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStop_Click(object sender, RoutedEventArgs e)
        {
            if (m != null)
            {
                m.Dispose();
                m = null;
            }
            this.btnStart.IsEnabled = true;
            this.btnStop.IsEnabled = false;
        }

        private void btnExit_Click(object sender, RoutedEventArgs e)
        {
            exitItem_Click(null, null);
        }

        /// <summary>
        /// 保存フォルダを開く
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReveal_Click(object sender, RoutedEventArgs e)
        {
            btnStop_Click(null, null); //記録停止

            try
            {
                string path = setting.SaveDir;
                System.Diagnostics.Process.Start(path);
            }
            catch (Exception) { }
        }

        /// <summary>
        /// 閉じるボタン
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Excel出力
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            btnStop_Click(null, null); //記録停止

            try
            {
                using (System.Windows.Forms.SaveFileDialog sfd = new System.Windows.Forms.SaveFileDialog())
                {
                    sfd.Filter = "Excel 97-2003 ブック(*.xls)|*.xls";
                    if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string template = setting.ExcelTemplate;
                        string save = sfd.FileName;
                        using (ExcelFile excel = new ExcelFile(template, save))
                        {
                            excel.Write(_captureList);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\n" + ex.StackTrace);
            }
        }


        private void listView1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            CaptureEntry entry = (CaptureEntry)listView1.SelectedItem;
            if (entry == null)
                return;

            if (entry.SelectedIndex == 1)
            {
                image1.Source = entry.ClipImage;
            }
            else
            {
                image1.Source = entry.WindowImage;
            }

        }

        private void listView1_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            listView1_SelectionChanged(null, null);
        }


    }
}
