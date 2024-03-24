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
            ApplicationConfiguration.Initialize();
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
    const int ICON_RED = 0;
    const int ICON_GREEN = 1;
    const int ICON_GRAY = 2;
    static int previousIconColor = ICON_GREEN;
    static string previousWindowTitle;
    static string previousprocessName;
    static bool ImeModeGlobal = true;
    static bool previousimeEnabled = true;
    static bool changeIme = false;
    static bool noKeyInput = false;
    static bool flagOnActivated = false;
    static bool ScheduleRunBackgroundApp = false;
    static int delayRunBackgroundApp = 2147483647;
    static string foregroundprocessName;
    private string foregroundWindowTitle;
    private readonly string[][] List = new string[5][];
    const int PassArray = 0;
    const int ImeOffArray = 1;
    const int ImeOffTitleArray = 2;
    const int OnActivatedAppArray = 3;
    const int BackgroundArray = 4;
    private string RunOnActivatedAppPath;
    private string RunOnActivatedArgv;
    private string RunBackgroundAppPath;
    private string FWBackgroundArgv;
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
        List[BackgroundArray] = null;
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
        buff = ConfigurationManager.AppSettings["DelayBackgroundTime"];
        if (!string.IsNullOrEmpty(buff))
        {
            DelayBackgroundInterval = (int)(float.Parse(buff) * 60 * 1000 / imeInterval);
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
                RunOnActivatedAppPath = buff;
            }
        }
        /* else
        {
            List[OnActivatedAppArray] = null;
        } */
        buff = ConfigurationManager.AppSettings["OnActivatedArgv"];
        if (!string.IsNullOrEmpty(buff))
        {
            RunOnActivatedArgv = buff;
        }
        buff = ConfigurationManager.AppSettings["BackgroundAppList"];
        if (!string.IsNullOrEmpty(buff))
        {
            List[BackgroundArray] = buff.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        }
        buff = (ConfigurationManager.AppSettings["BackgroundAppPath"]);
        if (!string.IsNullOrEmpty(buff))
        {
            if (!File.Exists(buff))
            {
                System.Windows.Forms.MessageBox.Show("BackgroundAppPath に指定したアプリが見つかりません");
                ScheduleRunBackgroundApp = false;
            }
            else
            {
                RunBackgroundAppPath = buff;
            }
        }
        /* else
        {
            List[BackgroundArray] = null;
        } */
        buff = ConfigurationManager.AppSettings["BackgroundArgv"];
        if (!string.IsNullOrEmpty(buff))
        {
            FWBackgroundArgv = buff;
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
            iconsize = (int)(16 * graphics.DpiX / 96);
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
        SetIcon(ICON_GREEN);
        Trace.WriteLine("クリックにより再開しました");
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
        SetIcon(ICON_GRAY);
        Trace.WriteLine($"{duration.TotalMinutes}分間無効にします({DateTime.Now + duration}まで)");

        this.timer.Stop();
        await Task.Delay(duration);
        this.timer.Start();

        SetIcon(ICON_GREEN);
        Trace.WriteLine("再開しました");
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
    private void SetIcon(int param)
    {
        if (ScheduleRunBackgroundApp && param == ICON_GREEN)
            param = ICON_RED;
        if (previousIconColor != param)
        {
            switch (param)
            {
                case ICON_RED:
                    icon.Icon = new Icon("Resources\\Red.ico", iconsize, iconsize);
                    break;
                case ICON_GRAY:
                    icon.Icon = new Icon("Resources\\Gray.ico", iconsize, iconsize);
                    break;
                default: /* ICON_GREEN */
                    icon.Icon = new Icon("Resources\\Green.ico", iconsize, iconsize);
                    break;
            }
        }
        previousIconColor = param;
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
                Debug.WriteLine($"{noKeyInputInterval / 1000}秒間キーボード入力がありません");
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
                Debug.WriteLine("IMEを有効にしました");
                changeIme = true;
            }
            else
            {
                Trace.WriteLine("IMEが有効になりません");
                changeIme = false;
            }
        }
        else
        {
            SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_SETOPENSTATUS, IntPtr.Zero);
            imeEnabled = (SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_GETOPENSTATUS, IntPtr.Zero) != 0);
            if (!imeEnabled)
            {
                Debug.WriteLine("IMEを無効にしました");
                changeIme = true;
            }
            else
            {
                Trace.WriteLine("IMEが無効になりません");
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
            Debug.WriteLine("SetImeOffList：IMEを無効にしました");
            changeIme = false;
            previousWindowTitle = foregroundWindowTitle;
        }
        else
        {
            Trace.WriteLine("SetImeOffList：IMEが無効になりません");
            changeIme = false;
        }
    }
    private void SetImePreset()
    {
        SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_SETOPENSTATUS, (IntPtr)1);
        imeEnabled = (SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_GETOPENSTATUS, IntPtr.Zero) != 0);
        if (imeEnabled)
        {
            Debug.WriteLine("IMEを有効にしました");
            changeIme = true;
        }
        else
        {
            Trace.WriteLine("IMEが有効になりません");
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
                    Trace.WriteLine($"{foregroundprocessName} により連携アプリが起動します");
                    Process.Start(RunOnActivatedAppPath, RunOnActivatedArgv);
                    flagOnActivated = true;
                    SetIcon(ICON_RED);
                }
            }
        }
    }
    private void MonitorBackground(int param)
    {
        Trace.WriteLine($"{previousprocessName},{foregroundprocessName}");
        // 非アクティブになってから指定回数ループで解除する
        if (List[param] != null)
        {
            delayRunBackgroundApp = DelayBackgroundInterval;
            SetIcon(ICON_RED);
            ScheduleRunBackgroundApp = true;
            Debug.WriteLine("Set ScheduleRunBackgroundApp = true;");
        }
    }
    private void RunBackgroundApp()
    {
        delayRunBackgroundApp--;
        if (delayRunBackgroundApp == 0)
        {
            // 何日かすると桁あふれするのでリセットする
            delayRunBackgroundApp = 2147483647;
            if (ScheduleRunBackgroundApp)
            {
                // 非アクティブならFirewallを解除する
                Process.Start(RunBackgroundAppPath, FWBackgroundArgv);
                ScheduleRunBackgroundApp = false;
                SetIcon(ICON_GREEN);
                Debug.WriteLine("時間経過により連携アプリが起動します");
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
            Trace.WriteLine("GetGUIThreadInfo failed");
            // スタートアップやロック解除時に例外0x80004005が発生する
            //throw new System.ComponentModel.Win32Exception();
            return;
        }

        imwd = ImmGetDefaultIMEWnd(gti.hwndFocus);

        // IMEの有効/無効の状態を確認
        imeConvMode = SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_GETCONVERSIONMODE, IntPtr.Zero);
        imeEnabled = (SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_GETOPENSTATUS, IntPtr.Zero) != 0);

        Debug.WriteLine($"{imeEnabled} status code:{imeConvMode}");
        if (!imeEnabled & imeConvMode == IME_CMODE_DISABLED)
        {
            Trace.WriteLine("IMEが無効です");
            changeIme = false;
            SetIcon(ICON_GREEN);
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
                            Debug.WriteLine($"タイトル:{foregroundWindowTitle} プロセス名:{foregroundprocessName}");
                        }
                        else
                        {
                            foregroundWindowTitle = "";
                            Trace.WriteLine($"タイトルを取得できません。タイトル:{foregroundWindowTitle} プロセス名:{foregroundprocessName}");
                            /* タイトルが空欄のプロセスが存在するので後でreturn;をする */
                            // return;
                        }
                    }
                    catch (Exception ex)
                    {
                        Trace.WriteLine("GetProcessById Failed: " + ex.Message);
                        return;
                    }
                }
                else
                {
                    Trace.WriteLine("GetWindowThreadProcessId Failed");
                    return;
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine("Getting window information Failed: " + ex.Message);
                return;
            }
        }
        else
        {
            Trace.WriteLine("GetForegroundWindow Failed");
            return;
        }

        if (CheckforegroundprocessName(PassArray))
        {
            Debug.WriteLine($"{foregroundprocessName} はPassListに含まれています");
            SetIcon(ICON_GRAY);
            return;
        }
        if (CheckforegroundprocessName(OnActivatedAppArray))
        {
            Debug.WriteLine($"{foregroundprocessName} はOnActivatedAppListに含まれています");
            MonitorOnActivated(OnActivatedAppArray);
        }
        else
        {
            flagOnActivated = false;
        }
        if (CheckpreviousprocessName(BackgroundArray))
        {
            Debug.WriteLine($"{previousprocessName} はBackgroundAppListに含まれています");
            MonitorBackground(BackgroundArray);
        }
        else
        {
            RunBackgroundApp();
        }
        if (string.IsNullOrEmpty(foregroundWindowTitle))
        {
            /* 先に出力してある */
            // Trace.WriteLine($"タイトルを取得できません。タイトル:{foregroundWindowTitle} プロセス名:{foregroundprocessName}");
            SetIcon(ICON_GREEN);
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
        SetIcon(ICON_GREEN);
    }
}
