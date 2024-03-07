using System;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Forms;

/*
 参考
 C#でIMEの変換モードを監視・変更する #C# - Qiita
 https://qiita.com/kob58im/items/a1644b36366f4d094a2c
 C#: タスクトレイに常駐するアプリの作り方 
 https://pineplanter.moo.jp/non-it-salaryman/2017/06/01/c-sharp-tasktray/
*/

class MainWindow
{
    static void Main()
    {
        bool createdNew;
        Mutex mutex = new Mutex(true, "AlwaysIME", out createdNew);

        if (createdNew)
        {
            ResidentTest rm = new ResidentTest();
            rm.InitializeAppArray();
            System.Windows.Forms.Application.Run();
            mutex.ReleaseMutex();
        }
    }
}

class ResidentTest : Form
{
    private System.Windows.Forms.Timer timer;
    static string previousWindowTitle;
    static bool previousimeEnabled = true;
    static bool changeIme = false;
    private string foregroundWindowTitle;
    private string[] appArray;
    private int imeInterval = 1000;

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("User32.dll")]
    static extern int SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

    [DllImport("imm32.dll")]
    static extern IntPtr ImmGetDefaultIMEWnd(IntPtr hWnd);

    [DllImport("user32.dll")]
    static extern bool GetGUIThreadInfo(uint dwthreadid, ref GUITHREADINFO lpguithreadinfo);

    [StructLayout(LayoutKind.Sequential)]
    public struct GUITHREADINFO
    {
        public int cbSize;
        public int flags;
        public IntPtr hwndActive;
        public IntPtr hwndFocus;
        public IntPtr hwndCapture;
        public IntPtr hwndMenuOwner;
        public IntPtr hwndMoveSize;
        public IntPtr hwndCaret;
        public System.Drawing.Rectangle rcCaret;
    }

    const int WM_IME_CONTROL = 0x283;
    const int IMC_GETCONVERSIONMODE = 1;
    const int IMC_SETCONVERSIONMODE = 2;
    const int IMC_GETOPENSTATUS = 5;
    const int IMC_SETOPENSTATUS = 6;

    const int IME_CMODE_DISABLED = 0;
    const int IME_CMODE_NATIVE = 1;
    const int IME_CMODE_KATAKANA = 2;
    const int IME_CMODE_FULLSHAPE = 8;
    const int IME_CMODE_ROMAN = 16;

    const int CMode_HankakuKana = IME_CMODE_ROMAN | IME_CMODE_KATAKANA | IME_CMODE_NATIVE;
    const int CMode_ZenkakuEisu = IME_CMODE_ROMAN | IME_CMODE_FULLSHAPE;
    const int CMode_Hiragana = IME_CMODE_ROMAN | IME_CMODE_FULLSHAPE | IME_CMODE_NATIVE;
    const int CMode_ZenkakuKana = IME_CMODE_ROMAN | IME_CMODE_FULLSHAPE | IME_CMODE_KATAKANA | IME_CMODE_NATIVE;
    // 実験してみた結果
    // 00 :× IMEが無効です                0000 0000
    // 19 :カ 半角カナ                     0001 0011
    // 24 :Ａ 全角英数                     0001 1000
    // 25 :あ ひらがな（漢字変換モード）   0001 1001
    // 27 :   全角カナ                     0001 1011

    public ResidentTest()
    {
        this.ShowInTaskbar = false;
        this.setComponents();
        this.timer = new System.Windows.Forms.Timer();
        this.timer.Interval = imeInterval;
        this.timer.Tick += new EventHandler(Timer_Tick);
        this.timer.Start();
    }

    public void InitializeAppArray()
    {
        string appList = ConfigurationManager.AppSettings["AppList"];
        if (!string.IsNullOrEmpty(appList))
        {
            appArray = appList.Split(',');
        }
        else
        {
            System.Windows.Forms.MessageBox.Show("AlwaysIME.exe.Config に異常があります。再インストールしてください。");
        }
        try
        {
            imeInterval = int.Parse(ConfigurationManager.AppSettings["intervalTime"]);
        }
        catch (Exception)
        {
            System.Windows.Forms.MessageBox.Show("AlwaysIME.exe.Config に異常があります。再インストールしてください。");
        }
    }

    private void Close_Click(object sender, EventArgs e)
    {
        System.Windows.Forms.Application.Exit();
    }
    private void setComponents()
    {
        float dpiX;
        using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
            dpiX = graphics.DpiX;
        NotifyIcon icon = new NotifyIcon();
        icon.Icon = new Icon("app.ico", (int)(SystemParameters.SmallIconWidth * dpiX / 96),
                                        (int)(SystemParameters.SmallIconHeight * dpiX / 96));
        icon.Visible = true;
        icon.Text = "AlwaysIME";
        ContextMenuStrip menu = new ContextMenuStrip();

        ToolStripMenuItem menuItemTop = new ToolStripMenuItem();
        menuItemTop.Text = "更新間隔";
        menu.Items.Add(menuItemTop);
        ToolStripMenuItem menuItem500 = new ToolStripMenuItem();
        menuItem500.Text = "> 500 ms";
        menuItem500.Click += new EventHandler((sender, e) => ChangeIntervalAndSave(500));
        menu.Items.Add(menuItem500);
        ToolStripMenuItem menuItem1000 = new ToolStripMenuItem();
        menuItem1000.Text = "> 1000 ms";
        menuItem1000.Click += new EventHandler((sender, e) => ChangeIntervalAndSave(1000));
        menu.Items.Add(menuItem1000);
        ToolStripMenuItem menuItem2000 = new ToolStripMenuItem();
        menuItem2000.Text = "> 2000 ms";
        menuItem2000.Click += new EventHandler((sender, e) => ChangeIntervalAndSave(2000));
        menu.Items.Add(menuItem2000);
        ToolStripMenuItem menuItem = new ToolStripMenuItem();
        menuItem.Text = "&終了";
        menuItem.Click += new EventHandler(Close_Click);
        menu.Items.Add(menuItem);
        icon.ContextMenuStrip = menu;
    }
    private void ChangeIntervalAndSave(int interval)
    {
        // App.config の更新
        Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        config.AppSettings.Settings["intervalTime"].Value = interval.ToString();
        config.Save(ConfigurationSaveMode.Modified);
        ConfigurationManager.RefreshSection("appSettings");

        // タイマーのインターバルを変更
        this.timer.Interval = interval;
    }
    private void Timer_Tick(object sender, EventArgs e)
    {
        MonitorActiveWindow();
    }

    // 半角カナ/全角英数/カタカナ モードを強制的に「ひらがな」モードに変更する
    void MonitorActiveWindow()
    {
        // IME状態の取得
        GUITHREADINFO gti = new GUITHREADINFO();
        gti.cbSize = Marshal.SizeOf(gti);

        if (!GetGUIThreadInfo(0, ref gti))
        {
            Console.WriteLine("GetGUIThreadInfo failed");
            // ロック解除時に例外0x80004005が発生する
            // throw new System.ComponentModel.Win32Exception();
            return;
        }

        IntPtr imwd = ImmGetDefaultIMEWnd(gti.hwndFocus);

        // IMEの有効/無効の状態を確認
        int imeConvMode = SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_GETCONVERSIONMODE, IntPtr.Zero);
        bool imeEnabled = (SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_GETOPENSTATUS, IntPtr.Zero) != 0);

#if DEBUG
        Console.WriteLine(imeEnabled.ToString() + " status code:" + imeConvMode.ToString());
#endif

        // アクティブウィンドウが変更された場合、IMEの状態を復元する
        IntPtr foregroundWindowHandle = GetForegroundWindow();
        const int nChars = 256;
        var buff = new System.Text.StringBuilder(nChars);
        if (GetWindowText(foregroundWindowHandle, buff, nChars) > 0)
        {
#if DEBUG
            Console.WriteLine("ウィンドウのタイトル: " + buff.ToString());
#endif
            foregroundWindowTitle = buff.ToString();
        }
        // プロセスIDを取得
        uint processId;
        GetWindowThreadProcessId(foregroundWindowHandle, out processId);
        // プロセスを取得
        Process process = Process.GetProcessById((int)processId);
        // プロセス名を取得
        string processName = process.ProcessName;
        Process[] processes = Process.GetProcessesByName(processName);
        // プロセス名を取得
        Process[] array = Process.GetProcesses();
        if (appArray != null)
        {
            for (int i = 0; i < appArray.Length; i++)
            {
                if (processName.ToLower() == appArray[i].ToLower())
                {
#if DEBUG
                    Console.WriteLine($"{processName} はAppListに含まれています。");
#endif
                    return;
                }
            }
        }
#if DEBUG
        Console.WriteLine("プロセス名: " + processName);
#endif
        if (foregroundWindowTitle != previousWindowTitle)
        {
            if (!imeEnabled & imeConvMode == IME_CMODE_DISABLED)
            {
                Console.WriteLine("IMEが無効です");
                changeIme = false;
            }
            else if (previousimeEnabled)
            {
                SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_SETOPENSTATUS, (IntPtr)1);
                imeEnabled = (SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_GETOPENSTATUS, IntPtr.Zero) != 0);
                if (imeEnabled)
                {
                    Console.WriteLine("IMEを有効にしました");
                    changeIme = true;
                }
                else
                {
                    Console.WriteLine("IMEが有効になりません");
                    changeIme = false;
                }
            }
            else
            {
                SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_SETOPENSTATUS, IntPtr.Zero);
                imeEnabled = (SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_GETOPENSTATUS, IntPtr.Zero) != 0);
                if (!imeEnabled)
                {
                    Console.WriteLine("IMEを無効にしました");
                    changeIme = true;
                }
                else
                {
                    Console.WriteLine("IMEが無効になりません");
                    changeIme = false;
                }
            }
        }
        if (imeEnabled)
        {
            switch (imeConvMode)
            {
                case CMode_Hiragana:
                    /* Nothing to do */
                    break;
                case CMode_HankakuKana: /* through */
                case CMode_ZenkakuEisu: /* through */
                case CMode_ZenkakuKana:
                    SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_SETCONVERSIONMODE, (IntPtr)CMode_Hiragana); // ひらがなモードに設定
                    break;
                default:
                    /* Nothing to do */
                    /* 環境によっては上のcaseをやめてここに飛ばしたほうがよいかも */
                    break;
            }
        }/* else 無変換(半角英数) */
        if (changeIme)
        {
            previousWindowTitle = foregroundWindowTitle;
            previousimeEnabled = imeEnabled;
        }
    }
}
