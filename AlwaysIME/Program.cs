using System;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
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
    private NotifyIcon icon;
    private int iconsize;
    static string previousWindowTitle;
    static string previousprocessName;
    static bool ImeModeGlobal = true;
    static bool previousimeEnabled = true;
    static bool changeIme = false;
    static bool noKeyInput = false;
    static bool flagFirewall = false;
    static bool flagRanappA = false;
    static bool flagRanappB = false;
    static int delayFirewall = 2147483647;
    static string foregroundprocessName;
    private string foregroundWindowTitle;
    private string[] appArray;
    private string[] ImeOffArray;
    private string[] FirewallBlockArrayA;
    private string[] FirewallBlockArrayB;
    private string FWappPathA;
    private string FWargvA;
    private string FWappPathB;
    private string FWargvB;
    private int imeInterval = 1000;
    private int SuspendFewInterval = 5;
    private int SuspendInterval = 45;
    private int FWAllowInterval = 2147483647;
    private int noKeyInputInterval = 6000;
    private DateTime lastInputTime;
    IntPtr imwd;
    int imeConvMode = 0;
    bool imeEnabled = false;

    [DllImport("user32.dll")]
    private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

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
    struct LASTINPUTINFO
    {
        public uint cbSize;
        public uint dwTime;
    }
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
        this.lastInputTime = DateTime.Now;
    }

    public void InitializeAppArray()
    {
        string AlwaysIMEMode = ConfigurationManager.AppSettings["AlwaysIMEMode"];
        if (!string.IsNullOrEmpty(AlwaysIMEMode))
        {
            if (AlwaysIMEMode.ToLower().CompareTo("off") == 0)
            {
                ImeModeGlobal = true;
            }
            else
            {
                ImeModeGlobal = false;
            }
        }
        else
        {
            ImeModeGlobal = false;
        }
        string appList = ConfigurationManager.AppSettings["AppList"];
        if (!string.IsNullOrEmpty(appList))
        {
            appArray = appList.Split(',');
        }
        else
        {
            /* Nothing to do */
        }
        string ImeOffList = ConfigurationManager.AppSettings["ImeOffList"];
        if (!string.IsNullOrEmpty(ImeOffList))
        {
            ImeOffArray = ImeOffList.Split(',');
        }
        else
        {
            /* Nothing to do */
        }
        try
        {
            imeInterval = int.Parse(ConfigurationManager.AppSettings["intervalTime"]);
        }
        catch (Exception)
        {
            /* Nothing to do */
        }
        try
        {
            SuspendFewInterval = int.Parse(ConfigurationManager.AppSettings["SuspendFewTime"]);
        }
        catch (Exception)
        {
            /* Nothing to do */
        }
        try
        {
            SuspendInterval = int.Parse(ConfigurationManager.AppSettings["SuspendTime"]);
        }
        catch (Exception)
        {
            /* Nothing to do */
        }
        try
        {
            noKeyInputInterval = (int)(float.Parse(ConfigurationManager.AppSettings["NoKeyInputTime"]) * 60 * 1000);
        }
        catch (Exception)
        {
            /* Nothing to do */
        }
        try
        {
            FWAllowInterval = int.Parse(ConfigurationManager.AppSettings["FWAllowCount"]);
        }
        catch (Exception)
        {
            /* Nothing to do */
        }
        string FirewallBlockListA = ConfigurationManager.AppSettings["FWBlockListA"];
        if (!string.IsNullOrEmpty(FirewallBlockListA))
        {
            FirewallBlockArrayA = FirewallBlockListA.Split(',');
        }
        else
        {
            /* Nothing to do */
        }
        FWappPathA = (ConfigurationManager.AppSettings["appPathA"]);
        if (!string.IsNullOrEmpty(FWappPathA))
        {
            if (!File.Exists(FWappPathA))
            {
                System.Windows.Forms.MessageBox.Show("FWappPathA に指定したアプリが見つかりません。");
                // Firewall機能を無効にする
                flagRanappA = false;
            }
        }
        /* else
        {
            FirewallBlockArrayA = null;
        } */
        try
        {
            FWargvA = (ConfigurationManager.AppSettings["argvA"]);
        }
        catch (Exception)
        {
            /* Nothing to do */
        }
        string FirewallBlockListB = ConfigurationManager.AppSettings["FWBlockListB"];
        if (!string.IsNullOrEmpty(FirewallBlockListB))
        {
            FirewallBlockArrayB = FirewallBlockListB.Split(',');
        }
        else
        {
            /* Nothing to do */
        }
        FWappPathB = (ConfigurationManager.AppSettings["appPathB"]);
        if (!string.IsNullOrEmpty(FWappPathB))
        {
            if (!File.Exists(FWappPathB))
            {
                System.Windows.Forms.MessageBox.Show("FWappPathB に指定したアプリが見つかりません。");
                flagRanappB = false;
            }
        }
        /* else
        {
            FirewallBlockArrayB = null;
        } */
        try
        {
            FWargvB = (ConfigurationManager.AppSettings["argvB"]);
        }
        catch (Exception)
        {
            /* Nothing to do */
        }
    }

    private void Close_Click(object sender, EventArgs e)
    {
        icon.Visible = false;
        icon.Dispose();
        System.Windows.Forms.Application.Exit();
    }
    private void setComponents()
    {
        using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
            iconsize = (int)(SystemParameters.SmallIconWidth * graphics.DpiX / 96);
        icon = new NotifyIcon();
        icon.Icon = new Icon("Resources\\Green.ico", iconsize, iconsize);
        icon.Visible = true;
        icon.Text = "AlwaysIME";
        ContextMenuStrip menu = new ContextMenuStrip();

        ToolStripMenuItem suspendFewMenuItem = new ToolStripMenuItem();
        suspendFewMenuItem.Text = "少し無効";
        suspendFewMenuItem.Click += new EventHandler(SuspendFewMenuItem_Click);
        menu.Items.Add(suspendFewMenuItem);
        ToolStripMenuItem suspendMenuItem = new ToolStripMenuItem();
        suspendMenuItem.Text = "しばらく無効";
        suspendMenuItem.Click += new EventHandler(SuspendMenuItem_Click);
        menu.Items.Add(suspendMenuItem);
        ToolStripMenuItem resumeMenuItem = new ToolStripMenuItem();
        resumeMenuItem.Text = "更新間隔";
        resumeMenuItem.Click += new EventHandler(ResumeMenuItem_Click);
        menu.Items.Add(resumeMenuItem);
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
    private void ResumeMenuItem_Click(object sender, EventArgs e)
    {
        this.timer.Start();
        icon.Icon = new Icon("Resources\\Green.ico", iconsize, iconsize);
        Console.WriteLine("クリックにより再開しました。");
    }
    private async void SuspendFewMenuItem_Click(object sender, EventArgs e)
    {
        await SuspendAsync(TimeSpan.FromMinutes(SuspendFewInterval));
    }
    private async void SuspendMenuItem_Click(object sender, EventArgs e)
    {
        await SuspendAsync(TimeSpan.FromMinutes(SuspendInterval));
    }

    private async Task SuspendAsync(TimeSpan duration)
    {
        icon.Icon = new Icon("Resources\\Gray.ico", iconsize, iconsize);
        Console.WriteLine($"{duration}分間無効にします。");

        this.timer.Stop();
        await Task.Delay(duration);
        this.timer.Start();

        icon.Icon = new Icon("Resources\\Green.ico", iconsize, iconsize);
        Console.WriteLine("再開しました。");
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
    private void SetIcon()
    {
        if (flagFirewall)
        {
            icon.Icon = new Icon("Resources\\Red.ico", iconsize, iconsize);
        }
        else
        {
            icon.Icon = new Icon("Resources\\Green.ico", iconsize, iconsize);
        }
    }
    private bool CheckProcessAppArray()
    {
        if (appArray != null)
        {
            for (int i = 0; i < appArray.Length; i++)
            {
                if (foregroundprocessName.ToLower() == appArray[i].ToLower())
                {
                    return true;
                }
            }
        }
        return false;
    }
    private bool CheckProcessImeOffArray()
    {
        if (ImeOffArray != null)
        {
            for (int i = 0; i < ImeOffArray.Length; i++)
            {
                if (foregroundprocessName.ToLower() == ImeOffArray[i].ToLower())
                {
                    return true;
                }
            }
        }
        return false;
    }
    private bool CheckProcessFirewallBlockArrayA()
    {
        if (FirewallBlockArrayA != null)
        {
            if (!string.IsNullOrEmpty(foregroundprocessName))
            {
                for (int i = 0; i < FirewallBlockArrayA.Length; i++)
                {
                    if (foregroundprocessName.ToLower() == FirewallBlockArrayA[i].ToLower())
                    {
                        return true;
                    }
                }
            }
        }
        return false;
    }
    private bool CheckProcessFirewallBlockArrayB()
    {
        if (FirewallBlockArrayB != null)
        {
            if (!string.IsNullOrEmpty(previousprocessName))
            {
                for (int i = 0; i < FirewallBlockArrayB.Length; i++)
                {
                    if (previousprocessName.ToLower() == FirewallBlockArrayB[i].ToLower())
                    {
                        return true;
                    }
                }
            }
        }
        return false;
    }
    private void CheckLastKeyInput()
    {
        LASTINPUTINFO lastInputInfo = new LASTINPUTINFO();
        lastInputInfo.cbSize = (uint)Marshal.SizeOf(lastInputInfo);
        GetLastInputInfo(ref lastInputInfo);
        lastInputTime = DateTime.Now.AddMilliseconds(-(Environment.TickCount - lastInputInfo.dwTime));
        if ((DateTime.Now - lastInputTime).TotalMilliseconds >= noKeyInputInterval)
        {
            if (!noKeyInput)
            {
#if DEBUG
                Console.WriteLine($"{noKeyInputInterval / 1000}秒間キーボード入力がありません");
#endif
                if (CheckProcessImeOffArray())
                {
                    SetImeOffList();
                }
                else
                {
                    SetImePreset();
                }
                noKeyInput = true;
            }
        }
        else
        {
            noKeyInput = false;
        }
    }

    private void SetImeGlobal()
    {
        if (previousimeEnabled)
        {
            SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_SETOPENSTATUS, (IntPtr)1);
            imeEnabled = (SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_GETOPENSTATUS, IntPtr.Zero) != 0);
            if (imeEnabled)
            {
#if DEBUG
                Console.WriteLine("IMEを有効にしました");
#endif
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
#if DEBUG
                Console.WriteLine("IMEを無効にしました");
#endif
                changeIme = true;
            }
            else
            {
                Console.WriteLine("IMEが無効になりません");
                changeIme = false;
            }
        }
    }

    private void SetImeOffList()
    {
        // previousWindowTitle：更新 previousimeEnabled：そのまま
        SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_SETOPENSTATUS, IntPtr.Zero);
        imeEnabled = (SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_GETOPENSTATUS, IntPtr.Zero) != 0);
        if (!imeEnabled)
        {
#if DEBUG
            Console.WriteLine("SetImeOffList：IMEを無効にしました");
#endif
            changeIme = false;
            previousWindowTitle = foregroundWindowTitle;
        }
        else
        {
            Console.WriteLine("SetImeOffList：IMEが無効になりません");
            changeIme = false;
        }
    }
    private void SetImePreset()
    {
        SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_SETOPENSTATUS, (IntPtr)1);
        imeEnabled = (SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_GETOPENSTATUS, IntPtr.Zero) != 0);
        if (imeEnabled)
        {
#if DEBUG
            Console.WriteLine("IMEを有効にしました");
#endif
            changeIme = true;
        }
        else
        {
            Console.WriteLine("IMEが有効になりません");
            changeIme = false;
        }
    }
    private void MonitorWindowA()
    {
        if (FirewallBlockArrayA != null)
        {
            for (int i = 0; i < FirewallBlockArrayA.Length; i++)
            {
                if (!string.IsNullOrEmpty(foregroundprocessName))
                {
                    if (foregroundprocessName.ToLower() == FirewallBlockArrayA[i].ToLower() & !flagRanappA)
                    {
                        // アクティブならFirewallでブロックする
                        Console.WriteLine($"{foregroundprocessName} によりFirewallをブロックします。");
                        Process.Start(FWappPathA, FWargvA);
                        flagRanappA = true;
                        flagFirewall = true;
                    }
                    else if (foregroundprocessName.ToLower() == FirewallBlockArrayA[i].ToLower() & flagRanappA)
                    {
                        flagRanappA = true;
                    }
                }
            }
        }
    }
    private void MonitorWindowB()
    {
        Console.WriteLine($"{previousprocessName},{foregroundprocessName}");
        // 非アクティブになってから{FWAllowCount}回ループで解除する
        if (FirewallBlockArrayB != null)
        {
            delayFirewall = FWAllowInterval;
            flagFirewall = true;
            flagRanappB = true;
#if DEBUG
            Console.WriteLine("Set flagRanappB = true;");
#endif
        }
    }
    private void RanappB()
    {
        delayFirewall--;
        if (delayFirewall == 0)
        {
            // 何日かすると桁あふれするのでリセットする
            delayFirewall = 2147483647;
            if (flagRanappB)
            {
                // 非アクティブならFirewallを解除する
                Process.Start(FWappPathB, FWargvB);
                flagFirewall = false;
                flagRanappB = false;
                icon.Icon = new Icon("Resources\\Green.ico", iconsize, iconsize);
#if DEBUG
                Console.WriteLine("時間経過によりFirewallを許可します");
#endif
            }
        }
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
            // スタートアップやロック解除時に例外0x80004005が発生する
            //throw new System.ComponentModel.Win32Exception();
            return;
        }

        imwd = ImmGetDefaultIMEWnd(gti.hwndFocus);

        // IMEの有効/無効の状態を確認
        imeConvMode = SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_GETCONVERSIONMODE, IntPtr.Zero);
        imeEnabled = (SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_GETOPENSTATUS, IntPtr.Zero) != 0);

#if DEBUG
        Console.WriteLine($"{imeEnabled} status code:{imeConvMode}");
#endif
        if (!imeEnabled & imeConvMode == IME_CMODE_DISABLED)
        {
            Console.WriteLine("IMEが無効です");
            changeIme = false;
            return;
        }

        // キーボード入力されたかチェック
        CheckLastKeyInput();

        previousprocessName = foregroundprocessName;
        // アクティブウィンドウが変更された場合、IMEの状態を復元する
        IntPtr foregroundWindowHandle = GetForegroundWindow();
        const int nChars = 256;
        var buff = new System.Text.StringBuilder(nChars);
        // プロセスIDを取得
        uint processId;
        GetWindowThreadProcessId(foregroundWindowHandle, out processId);
        // プロセスを取得
        Process process = Process.GetProcessById((int)processId);
        // プロセス名を取得
        foregroundprocessName = process.ProcessName;
        Process[] processes = Process.GetProcessesByName(foregroundprocessName);
        // プロセス名を取得
        Process[] array = Process.GetProcesses();

        if (GetWindowText(foregroundWindowHandle, buff, nChars) > 0)
        {
#if DEBUG
            Console.WriteLine($"タイトル:{buff} プロセス名:{foregroundprocessName}");
#endif
            foregroundWindowTitle = buff.ToString();
        }
        else
        {
            Console.WriteLine($"タイトルを取得できません。タイトル:{buff} プロセス名:{foregroundprocessName}");
            return;
        }

        if (CheckProcessAppArray())
        {
#if DEBUG
            Console.WriteLine($"{foregroundprocessName} はAppListに含まれています。");
#endif
            icon.Icon = new Icon("Resources\\Gray.ico", iconsize, iconsize);
            return;
        }
        if (CheckProcessFirewallBlockArrayA())
        {
#if DEBUG
            Console.WriteLine($"{foregroundprocessName} はFirewallBlockArrayAに含まれています。");
#endif
            MonitorWindowA();
        }
        if (CheckProcessFirewallBlockArrayB())
        {
#if DEBUG
            Console.WriteLine($"{previousprocessName} はFirewallBlockArrayBに含まれています。");
#endif
            MonitorWindowB();
        }
        else
        {
            RanappB();
        }
        if (foregroundWindowTitle != previousWindowTitle)
        {
            if (ImeModeGlobal)
            {
                if (CheckProcessImeOffArray())
                {
                    SetImeOffList();
                }
                else
                {
                    SetImeGlobal();
                }
            }
            if (!ImeModeGlobal)
            {
                if (CheckProcessImeOffArray())
                {
                    SetImeOffList();
                }
                else
                {
                    SetImePreset();
                }
            }
            SetIcon();
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
