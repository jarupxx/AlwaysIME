using Microsoft.Win32;
using System;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
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
            rm.InitializeAppConfig();
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
    static bool flagIconColor = false;
    static bool flagOnActivated = false;
    static bool ScheduleRanEnteredBackgroundApp = false;
    static int delayRanEnteredBackgroundApp = 2147483647;
    static string foregroundprocessName;
    private string foregroundWindowTitle;
    private readonly string[][] List = new string[6][];
    const int PassArray = 0;
    const int ImeOffArray = 1;
    const int ImeOffTitleArray = 2;
    const int OnActivatedAppArray = 4;
    const int EnteredBackgroundArray = 5;
    private string RanOnActivatedAppPath;
    private string RanOnActivatedArgv;
    private string RanEnteredBackgroundAppPath;
    private string FWEnteredBackgroundArgv;
    private int imeInterval = 500;
    private int SuspendFewInterval = 5;
    private int SuspendInterval = 45;
    private int DelayBackgroundInterval = 2147483647;
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

    public void InitializeAppConfig()
    {
        List[PassArray] = null;
        List[ImeOffArray] = null;
        List[ImeOffTitleArray] = null;
        List[OnActivatedAppArray] = null;
        List[EnteredBackgroundArray] = null;
        string buff = ConfigurationManager.AppSettings["AlwaysIMEMode"];
        if (!string.IsNullOrEmpty(buff))
        {
            if (buff.ToLower().CompareTo("off") == 0)
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
        buff = ConfigurationManager.AppSettings["PassList"];
        if (!string.IsNullOrEmpty(buff))
        {
            List[PassArray] = buff.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        }
        buff = ConfigurationManager.AppSettings["ImeOffList"];
        if (!string.IsNullOrEmpty(buff))
        {
            List[ImeOffArray] = buff.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        }
        buff = ConfigurationManager.AppSettings["ImeOffTitle"];
        if (!string.IsNullOrEmpty(buff))
        {
            List[ImeOffTitleArray] = buff.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        }
        buff = ConfigurationManager.AppSettings["intervalTime"];
        if (!string.IsNullOrEmpty(buff))
        {
            imeInterval = int.Parse(buff);
        }
        buff = ConfigurationManager.AppSettings["SuspendFewTime"];
        if (!string.IsNullOrEmpty(buff))
        {
            SuspendFewInterval = int.Parse(buff);
        }
        buff = ConfigurationManager.AppSettings["SuspendTime"];
        if (!string.IsNullOrEmpty(buff))
        {
            SuspendInterval = int.Parse(buff);
        }
        buff = ConfigurationManager.AppSettings["NoKeyInputTime"];
        if (!string.IsNullOrEmpty(buff))
        {
            noKeyInputInterval = (int)(float.Parse(buff) * 60 * 1000);
        }
        buff = ConfigurationManager.AppSettings["DelayBackgroundCount"];
        if (!string.IsNullOrEmpty(buff))
        {
            DelayBackgroundInterval = int.Parse(buff);
        }
        buff = ConfigurationManager.AppSettings["OnActivatedAppList"];
        if (!string.IsNullOrEmpty(buff))
        {
            List[OnActivatedAppArray] = buff.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        }
        buff = (ConfigurationManager.AppSettings["OnActivatedAppPath"]);
        if (!string.IsNullOrEmpty(buff))
        {
            if (!File.Exists(buff))
            {
                System.Windows.Forms.MessageBox.Show("OnActivatedAppPath に指定したアプリが見つかりません");
            }
            else
            {
                RanOnActivatedAppPath = buff;
            }
        }
        /* else
        {
            List[OnActivatedAppArray] = null;
        } */
        buff = ConfigurationManager.AppSettings["OnActivatedArgv"];
        if (!string.IsNullOrEmpty(buff))
        {
            RanOnActivatedArgv = buff;
        }
        buff = ConfigurationManager.AppSettings["EnteredBackgroundAppList"];
        if (!string.IsNullOrEmpty(buff))
        {
            List[EnteredBackgroundArray] = buff.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        }
        buff = (ConfigurationManager.AppSettings["EnteredBackgroundAppPath"]);
        if (!string.IsNullOrEmpty(buff))
        {
            if (!File.Exists(buff))
            {
                System.Windows.Forms.MessageBox.Show("EnteredBackgroundAppPath に指定したアプリが見つかりません");
                ScheduleRanEnteredBackgroundApp = false;
            }
            else
            {
                RanEnteredBackgroundAppPath = buff;
            }
        }
        /* else
        {
            List[EnteredBackgroundArray] = null;
        } */
        buff = ConfigurationManager.AppSettings["EnteredBackgroundArgv"];
        if (!string.IsNullOrEmpty(buff))
        {
            FWEnteredBackgroundArgv = buff;
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
        suspendFewMenuItem.Text = "少し無効(&S)";
        suspendFewMenuItem.Click += new EventHandler(SuspendFewMenuItem_Click);
        menu.Items.Add(suspendFewMenuItem);
        ToolStripMenuItem suspendMenuItem = new ToolStripMenuItem();
        suspendMenuItem.Text = "しばらく無効(&P)";
        suspendMenuItem.Click += new EventHandler(SuspendMenuItem_Click);
        menu.Items.Add(suspendMenuItem);
        ToolStripMenuItem resumeMenuItem = new ToolStripMenuItem();
        resumeMenuItem.Text = "再度有効(&R)";
        resumeMenuItem.Click += new EventHandler(ResumeMenuItem_Click);
        menu.Items.Add(resumeMenuItem);

        ToolStripSeparator separator = new ToolStripSeparator();
        menu.Items.Add(separator);

        ToolStripMenuItem updateTimeMenuItem = new ToolStripMenuItem("更新時間");
        ToolStripMenuItem menuItem250 = new ToolStripMenuItem("250 ms");
        menuItem250.Click += new EventHandler((sender, e) => ChangeIntervalAndSave(250));
        ToolStripMenuItem menuItem500 = new ToolStripMenuItem("500 ms");
        menuItem500.Click += new EventHandler((sender, e) => ChangeIntervalAndSave(500));
        ToolStripMenuItem menuItem1000 = new ToolStripMenuItem("1000 ms");
        menuItem1000.Click += new EventHandler((sender, e) => ChangeIntervalAndSave(1000));
        updateTimeMenuItem.DropDownItems.Add(menuItem250);
        updateTimeMenuItem.DropDownItems.Add(menuItem500);
        updateTimeMenuItem.DropDownItems.Add(menuItem1000);
        menu.Items.Add(updateTimeMenuItem);

        ToolStripSeparator separator2 = new ToolStripSeparator();
        menu.Items.Add(separator2);

        ToolStripMenuItem menuItem = new ToolStripMenuItem();
        menuItem.Text = "常駐の終了(&X)";
        menuItem.Click += new EventHandler(Close_Click);
        menu.Items.Add(menuItem);
        icon.ContextMenuStrip = menu;
    }
    private void ResumeMenuItem_Click(object sender, EventArgs e)
    {
        this.timer.Start();
        icon.Icon = new Icon("Resources\\Green.ico", iconsize, iconsize);
        Console.WriteLine("クリックにより再開しました");
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
        Console.WriteLine($"{duration.TotalMinutes}分間無効にします({DateTime.Now + duration}まで)");

        this.timer.Stop();
        await Task.Delay(duration);
        this.timer.Start();

        icon.Icon = new Icon("Resources\\Green.ico", iconsize, iconsize);
        Console.WriteLine("再開しました");
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
        if (flagIconColor)
        {
            icon.Icon = new Icon("Resources\\Red.ico", iconsize, iconsize);
        }
        else
        {
            icon.Icon = new Icon("Resources\\Green.ico", iconsize, iconsize);
        }
    }
    private bool CheckforegroundprocessName(int param)
    {
        if (List[param] != null)
        {
            if (!string.IsNullOrEmpty(foregroundprocessName))
            {
                for (int i = 0; i < List[param].Length; i++)
                {
                    if (foregroundprocessName.ToLower() == List[param][i].ToLower())
                    {
                        return true;
                    }
                }
            }
        }
        return false;
    }
    private bool CheckforegroundWindowTitleRegex(int param)
    {
        if (List[param] != null)
        {
            for (int i = 0; i < List[param].Length; i++)
            {
                if (Regex.IsMatch(foregroundWindowTitle, List[param][i]))
                {
                    return true;
                }
            }
        }
        return false;
    }
    private bool CheckpreviousprocessName(int param)
    {
        if (List[param] != null)
        {
            if (!string.IsNullOrEmpty(previousprocessName))
            {
                for (int i = 0; i < List[param].Length; i++)
                {
                    if (previousprocessName.ToLower() == List[param][i].ToLower())
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
                if (CheckforegroundprocessName(ImeOffArray))
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
    private void MonitorOnActivated(int param)
    {
        if (List[param] != null)
        {
            if (!string.IsNullOrEmpty(foregroundprocessName))
            {
                if (!flagOnActivated)
                {
                    Console.WriteLine($"{foregroundprocessName} により連携アプリが起動します");
                    Process.Start(RanOnActivatedAppPath, RanOnActivatedArgv);
                    flagOnActivated = true;
                    flagIconColor = true;
                    icon.Icon = new Icon("Resources\\Red.ico", iconsize, iconsize);
                }
            }
        }
    }
    private void MonitorEnteredBackground(int param)
    {
        Console.WriteLine($"{previousprocessName},{foregroundprocessName}");
        // 非アクティブになってから{DelayBackgroundCount}回ループで解除する
        if (List[param] != null)
        {
            delayRanEnteredBackgroundApp = DelayBackgroundInterval;
            flagIconColor = true;
            ScheduleRanEnteredBackgroundApp = true;
#if DEBUG
            Console.WriteLine("Set ScheduleRanEnteredBackgroundApp = true;");
#endif
        }
    }
    private void RanEnteredBackgroundApp()
    {
        delayRanEnteredBackgroundApp--;
        if (delayRanEnteredBackgroundApp == 0)
        {
            // 何日かすると桁あふれするのでリセットする
            delayRanEnteredBackgroundApp = 2147483647;
            if (ScheduleRanEnteredBackgroundApp)
            {
                // 非アクティブならFirewallを解除する
                Process.Start(RanEnteredBackgroundAppPath, FWEnteredBackgroundArgv);
                flagIconColor = false;
                ScheduleRanEnteredBackgroundApp = false;
                icon.Icon = new Icon("Resources\\Green.ico", iconsize, iconsize);
#if DEBUG
                Console.WriteLine("時間経過により連携アプリが起動します");
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

        // キーボード未入力ならIMEの状態を復元する
        CheckLastKeyInput();

        // プロセスの変更を追跡するために保存する
        previousprocessName = foregroundprocessName;
        // アクティブウィンドウが変更された場合、IMEの状態を復元する
        IntPtr foregroundWindowHandle = GetForegroundWindow();
        if (foregroundWindowHandle != IntPtr.Zero)
        {
            try
            {
                // プロセスIDを取得
                if (GetWindowThreadProcessId(foregroundWindowHandle, out uint processId) != 0)
                {
                    try
                    {
                        // プロセスを取得
                        Process process = Process.GetProcessById((int)processId);
                        // プロセス名を取得
                        foregroundprocessName = process.ProcessName;
                        // ウィンドウのタイトルを取得
                        const int nChars = 256;
                        StringBuilder titleBuilder = new StringBuilder(nChars);
                        if (GetWindowText(foregroundWindowHandle, titleBuilder, titleBuilder.Capacity) > 0)
                        {
                            foregroundWindowTitle = titleBuilder.ToString();
#if DEBUG
                            Console.WriteLine($"タイトル:{foregroundWindowTitle} プロセス名:{foregroundprocessName}");
#endif
                        }
                        else
                        {
                            foregroundWindowTitle = "";
                            Console.WriteLine($"タイトルを取得できません。タイトル:{foregroundWindowTitle} プロセス名:{foregroundprocessName}");
                            /* タイトルが空欄のプロセスが存在するので後でreturn;をする */
                            // return;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("GetProcessById Failed: " + ex.Message);
                        return;
                    }
                }
                else
                {
                    Console.WriteLine("GetWindowThreadProcessId Failed");
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Getting window information Failed: " + ex.Message);
                return;
            }
        }
        else
        {
            Console.WriteLine("GetForegroundWindow Failed");
            return;
        }

        if (CheckforegroundprocessName(PassArray))
        {
#if DEBUG
            Console.WriteLine($"{foregroundprocessName} はAppListに含まれています");
#endif
            icon.Icon = new Icon("Resources\\Gray.ico", iconsize, iconsize);
            return;
        }
        if (CheckforegroundprocessName(OnActivatedAppArray))
        {
#if DEBUG
            Console.WriteLine($"{foregroundprocessName} はOnActivatedAppListに含まれています");
#endif
            MonitorOnActivated(OnActivatedAppArray);
        }
        else
        {
            flagOnActivated = false;
        }
        if (CheckpreviousprocessName(EnteredBackgroundArray))
        {
#if DEBUG
            Console.WriteLine($"{previousprocessName} はEnteredBackgroundAppListに含まれています");
#endif
            MonitorEnteredBackground(EnteredBackgroundArray);
        }
        else
        {
            RanEnteredBackgroundApp();
        }
        if (string.IsNullOrEmpty(foregroundWindowTitle))
        {
            /* 先に出力してある */
            // Console.WriteLine($"タイトルを取得できません。タイトル:{foregroundWindowTitle} プロセス名:{foregroundprocessName}");
            return;
        }
        if (foregroundWindowTitle != previousWindowTitle)
        {
            if (ImeModeGlobal)
            {
                if (CheckforegroundprocessName(ImeOffArray))
                {
                    SetImeOffList();
                }
                else if (CheckforegroundWindowTitleRegex(ImeOffTitleArray))
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
                if (CheckforegroundprocessName(ImeOffArray))
                {
                    SetImeOffList();
                }
                else if (CheckforegroundWindowTitleRegex(ImeOffTitleArray))
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
