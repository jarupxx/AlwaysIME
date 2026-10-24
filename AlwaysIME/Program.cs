using Microsoft.Win32;
using System.Configuration;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

/*
 参考
 C#でIMEの変換モードを監視・変更する #C# - Qiita
 https://qiita.com/kob58im/items/a1644b36366f4d094a2c
 C#: タスクトレイに常駐するアプリの作り方
 https://pineplanter.moo.jp/non-it-salaryman/2017/06/01/c-sharp-tasktray/
*/

class MainWindow
{
    static Mutex mutex;
    static EventWaitHandle restartSignalEvent;
    const string RestartSignalEventName = "AlwaysIME_RestartSignal";
    static RegisteredWaitHandle restartSignalRegisteredWait;

    [STAThread]
    static void Main()
    {
        bool createdNew;
        mutex = new Mutex(true, "AlwaysIME", out createdNew);

        if (createdNew)
        {
            // フォームを作る前に、未処理例外の記録を有効にする
            GlobalExceptionHandler.Install();
            AppLog.LogStartup();
            ApplicationConfiguration.Initialize();
            // 設定ファイルの場所を決定する（ConfigurationManagerを最初に使う前に行う必要がある）
            if (!ConfigStore.Initialize())
            {
                AppLog.EnsureLocation();
                AppLog.Warn("設定ファイルの場所を決定できなかったため、起動を中止します");
                mutex.ReleaseMutex();
                return;
            }
            // ログの保存先を、設定ファイルと同じ方式（ポータブル / ユーザー別）で確定する
            AppLog.Configure(ConfigStore.IsPortable);
            AppLog.Info($"設定ファイル: {ConfigStore.FilePath}（{(ConfigStore.IsPortable ? "ポータブル" : "ユーザー別")}）");
            // 設定の検証・復元と、バックアップ監視の開始
            if (!BackupService.OnStartup())
            {
                AppLog.Warn("設定の検証・復元ができなかったため、起動を中止します");
                mutex.ReleaseMutex();
                return;
            }
            ResidentTest.InitializeAppConfig();
            ResidentTest rm = new ResidentTest();
            WatchForRestartSignal();
            // 実行ファイルを移動していた場合に、自動起動の登録を新しい場所へ更新する（起動を遅らせないよう別スレッドで行う）
            Task.Run(AutoStart.VerifyRegisteredPath);
            Application.Run();
            AppLog.Info("終了します");
            BackupService.OnExit();
            restartSignalRegisteredWait?.Unregister(null);
            restartSignalEvent?.Dispose();
            mutex.ReleaseMutex();
        }
        else
        {
            // すでに起動しているインスタンスがある場合は、再起動シグナルを送信
            // （ログの設定を読めるよう、既存の設定ファイルの場所を使う）
            ConfigStore.TryActivateExisting();
            AppLog.EnsureLocation();
            try
            {
                using EventWaitHandle signalEvent = EventWaitHandle.OpenExisting(RestartSignalEventName);
                signalEvent.Set();
                AppLog.Info("起動中のインスタンスへ再起動シグナルを送信しました");
            }
            catch (WaitHandleCannotBeOpenedException)
            {
                // 既存インスタンスがまだイベントを作成していない場合は何もできない
                AppLog.Warn("再起動シグナル用イベントが見つかりません。");
            }
        }
    }

    static void WatchForRestartSignal()
    {
        // AutoReset: Set()を受け取ったら自動的に非シグナル状態に戻る
        restartSignalEvent = new EventWaitHandle(false, EventResetMode.AutoReset, RestartSignalEventName);

        // イベントがシグナルされたらスレッドプールのスレッドでコールバックを実行する
        restartSignalRegisteredWait = ThreadPool.RegisterWaitForSingleObject(
            restartSignalEvent,
            (state, timedOut) => OnRestartSignalReceived(),
            null,
            Timeout.Infinite,
            executeOnlyOnce: false);
    }

    static void OnRestartSignalReceived()
    {
        RestartApplication();
    }

    static void RestartApplication()
    {
        AppLog.Info("再起動します");
        // 新しいプロセスを起動する前に、終了時のバックアップを済ませておく
        BackupService.OnExit();
        string exePath = Process.GetCurrentProcess().MainModule.FileName;
        Process.Start(exePath);

        // 現在のプロセスを終了
        Environment.Exit(0);
    }
}

class ResidentTest : Form
{
    private System.Windows.Forms.Timer timer;
    private NotifyIcon icon;
    private int iconsize;
    private ToolStripMenuItem menuPunctuation;
    private ToolStripMenuItem menuSpace;
    private ToolStripSeparator separatorPunct;
    private ToolStripSeparator separatorSpace;
    private static ResidentTest currentInstance;
    const int ICON_RED = 0;
    const int ICON_GREEN = 1;
    const int ICON_GRAY = 2;
    static int previousIconColor = ICON_GREEN;
    static string previousWindowTitle;
    static string previousprocessName;
    static string RegistrationprocessName;
    const int IMEMODE_ON = 1;
    const int IMEMODE_OFF = 2;
    const int IMEMODE_GLOBAL = 3;
    static int ImeModeGlobal = IMEMODE_ON;
    internal static bool darkModeEnabled = false;
    static bool previousimeEnabled = true;
    internal static bool AppExitEndEnabled = true;
    static bool changeIme = false;
    static bool noKeyInput = false;
    static bool flagOnActivated = false;
    static bool ScheduleRunBackgroundApp = false;
    static int delayRunBackgroundApp = 2147483647;
    static string foregroundprocessName;
    private static string foregroundWindowTitle;
    static string RegistrationWindowTitle;
    internal static readonly string[][] List = new string[6][];
    internal const int PassArray = 0;
    internal const int ImeOffArray = 1;
    internal const int ImeOffTitleArray = 2;
    internal const int RemoveUpdateTagArray = 3;
    const int OnActivatedAppArray = 4;
    const int BackgroundArray = 5;
    private static string RunOnActivatedAppPath;
    private static string RunOnActivatedArgv;
    private static string RunBackgroundAppPath;
    private static string FWBackgroundArgv;
    private static int imeInterval = 200;
    private static int SuspendFewInterval = 5;
    private static int SuspendInterval = 45;
    private static int DelayBackgroundInterval = 2147483647;
    internal static int noKeyInputInterval = 6000;
    private DateTime lastInputTime;
    IntPtr imwd;
    int imeConvMode = 0;
    bool imeEnabled = false;
    internal static readonly char[] separator = [','];

    [DllImport("user32.dll")]
    private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

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
        public Rectangle rcCaret;
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
    const int CModeMS_HankakuKana = IME_CMODE_KATAKANA | IME_CMODE_NATIVE;
    const int CModeMS_ZenkakuEisu = IME_CMODE_FULLSHAPE;
    const int CModeMS_Hiragana = IME_CMODE_FULLSHAPE | IME_CMODE_NATIVE;
    const int CModeMS_ZenkakuKana = IME_CMODE_FULLSHAPE | IME_CMODE_KATAKANA | IME_CMODE_NATIVE;
    // MS-IMEはIME_CMODE_ROMANが立たなくなった (2024/4)
    // 00 :× IMEが無効です                0000 0000
    // 03 :カ 半角カナ                     0000 0011
    // 08 :Ａ 全角英数                     0000 1000
    // 09 :あ ひらがな（漢字変換モード）   0000 1001
    // 11 :   全角カナ                     0000 1011

    static readonly int[][] val = new int[3][];
    const int ConfigPunctuation = 0;
    const int ConfigSpaceWidth = 1;
    static int SetPunctuationMode;
    static int SetSpaceMode;
    // MS-IME入力設定(レジストリ)
    const string MsImeRegistryPath = @"Software\Microsoft\IME\15.0\IMEJP\MSIME";
    // 変換/無変換キーへのIMEオン・オフ割り当て: 変更前の値を保持して終了時に復元する
    static bool keyAssignmentRecorded = false;
    static int origKeyAssignment = -1; // -1 = 元々値が存在しなかった
    static readonly string[] keyPath = new string[2];
    static readonly string[] valueName = new string[2];
    static readonly RegistryValueKind[] valueType = new RegistryValueKind[2];
    // 0: は「現在の入力モード」だが「ひらがな（漢字変換モード）」しか想定していないので全角スペースとする
    private readonly string[] SpaceWidthText = ["全角", "全角", "半角"];
    internal const int IME_AUTO_WIDTH_SPACE = 0;
    internal const int IME_FULL_WIDTH_SPACE = 1;
    internal const int IME_HALF_WIDTH_SPACE = 2;
    // 0: 現在の入力モード
    // 1: 常に全角
    // 2: 常に半角

    private readonly string[] PunctuationText = ["，．", "、。", "、．", "，。"];
    const int IME_COMMA_PERIOD = 0;
    internal const int IME_TOUTEN_KUTENN = 1;
    const int IME_TOUTEN_PERIOD = 2;
    internal const int IME_COMMA_KUTENN = 3;
    // 0001 01XX 0010 0000
    // ビットマスク(mode >> 16) & 0x3
    // 0001 0100 0010 0000	COMMA_PERIOD = 0
    // 0001 0101 0010 0000	TOUTEN_KUTENN = 1
    // 0001 0110 0010 0000	TOUTEN_PERIOD = 2
    // 0001 0111 0010 0000	COMMA_KUTENN = 3
    // 0: ，．
    // 1: 、。
    // 2: 、．
    // 3: ，。

    public ResidentTest()
    {
        this.ShowInTaskbar = false;
        currentInstance = this;
        this.setComponents();
        this.timer = new System.Windows.Forms.Timer();
        this.timer.Interval = imeInterval;
        this.timer.Tick += new EventHandler(Timer_Tick);
        this.timer.Start();
        this.lastInputTime = DateTime.Now;
    }

    // 設定項目ごとの初期値(AlwaysIME.dll.configへ書き戻す際に使用する)
    internal static readonly Dictionary<string, string> DefaultConfigValues = new Dictionary<string, string>
    {
        { "AlwaysIMEMode", "1" },
        { "intervalTime", "200" },
        { "SuspendFewTime", "5" },
        { "SuspendTime", "45" },
        { "NoKeyInputTime", "0.1" },
        { "DelayBackgroundTime", "" },
        { "Punctuation", "0,0" },
        { "SpaceWidth", "0,0" },
        { "AppExitEnd", "1" },
        { "MsImeSwitch", "0" },
        { "KeyAssignment", "" },
        { "IsDarkMode", "off" },
        { "PassList", "" },
        { "ImeOffList", "" },
        { "ImeOffTitle", "" },
        { "UpdateTag", "" },
        { "OnActivatedAppList", "" },
        { "OnActivatedAppPath", "" },
        { "OnActivatedArgv", "" },
        { "BackgroundAppList", "" },
        { "BackgroundAppPath", "" },
        { "BackgroundArgv", "" },
        { "BackupStableHours", "24" },
        { "BackupKeep", "5" },
        { "LogEnabled", "1" },
        { "LogLevel", "info" },
        { "LogMaxSizeMB", "1" },
        { "LogKeep", "3" },
    };

    // 設定値が不正な場合に確認ダイアログを表示する。
    // 「初期値にする」: 初期値をAlwaysIME.dll.configへ書き戻して続行(この呼び出し元の
    //                   フォールバック処理でメモリ上の値も初期値になる)。
    // 「設定を確認」  : AlwaysIME.dll.configを既定のアプリで開き、閉じたら設定を
    //                   再読み込みして最初からやり直す。
    private static void HandleInvalidConfigValue(string settingKey)
    {
        using (var dialog = new ConfigWarningDialog($"{settingKey}の設定が間違っています。"))
        {
            DialogResult result = dialog.ShowDialog();
            if (result == DialogResult.Yes)
            {
                // 初期値にする
                UpdateConfigDefaultValue(settingKey);
            }
            else
            {
                // 設定を確認する
                OpenConfigFileAndReload();
            }
        }
    }

    // 指定した設定項目の値をAlwaysIME.dll.config上で初期値に書き換える
    private static void UpdateConfigDefaultValue(string settingKey)
    {
        if (!DefaultConfigValues.TryGetValue(settingKey, out string defaultValue))
        {
            return;
        }

        try
        {
            var map = new ExeConfigurationFileMap { ExeConfigFilename = ConfigStore.FilePath };
            Configuration config = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);

            if (config.AppSettings.Settings[settingKey] != null)
            {
                config.AppSettings.Settings[settingKey].Value = defaultValue;
            }
            else
            {
                config.AppSettings.Settings.Add(settingKey, defaultValue);
            }

            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");
        }
        catch (Exception ex)
        {
            MessageBox.Show("設定ファイルへの書き込みに失敗しました: " + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // AlwaysIME.dll.configを既定のアプリケーションで開き、閉じられたら設定を
    // 再読み込みして最初からやり直す。再読み込みに成功した場合は
    // ConfigReloadRequestedExceptionを送出し、呼び出し元に残る古い状態の
    // 処理(不正値のままの続行部分)を打ち切る。
    private static void OpenConfigFileAndReload()
    {
        string configFilePath = ConfigStore.FilePath;

        try
        {
            Process editorProcess = Process.Start(new ProcessStartInfo
            {
                FileName = configFilePath,
                UseShellExecute = true
            });

            if (editorProcess != null)
            {
                // エディタが閉じられるまで待機
                editorProcess.WaitForExit();

                // 変更後の内容を反映して最初からやり直す
                ConfigurationManager.RefreshSection("appSettings");
                InitializeAppConfig();

                // 上のInitializeAppConfig()で最新状態まで処理済みのため、
                // 呼び出し元(古い設定値のまま処理を続けようとしている箇所)を打ち切る
                throw new ConfigReloadRequestedException();
            }
        }
        catch (ConfigReloadRequestedException)
        {
            throw;
        }
        catch (Exception ex)
        {
            MessageBox.Show("設定ファイルを開けませんでした: " + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // 設定値が不正なときに「初期値にする」か「設定を確認」かを選択させるダイアログ
    private class ConfigWarningDialog : Form
    {
        public ConfigWarningDialog(string message)
        {
            float UIScale;
            using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
                UIScale = graphics.DpiX / 96;
            this.Font = new Font("Meiryo", (int)(9 * Math.Pow(UIScale, 1.0 / 3.0)), FontStyle.Regular);
            this.Text = "AlwaysIME";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new Size((int)(400 * UIScale), (int)(170 * UIScale));

            Label messageLabel = new Label();
            messageLabel.Text = message + "\n初期値にするか、設定を確認してください。";
            messageLabel.Location = new Point((int)(20 * UIScale), (int)(20 * UIScale));
            messageLabel.Size = new Size((int)(360 * UIScale), (int)(60 * UIScale));

            Button defaultButton = new Button();
            defaultButton.Text = "初期値にする(&D)";
            defaultButton.Size = new Size((int)(160 * UIScale), (int)(32 * UIScale));
            defaultButton.Location = new Point((int)(30 * UIScale), (int)(95 * UIScale));
            defaultButton.DialogResult = DialogResult.Yes;

            Button checkButton = new Button();
            checkButton.Text = "設定を確認(&C)";
            checkButton.Size = new Size((int)(160 * UIScale), (int)(32 * UIScale));
            checkButton.Location = new Point((int)(205 * UIScale), (int)(95 * UIScale));
            checkButton.DialogResult = DialogResult.No;

            if (darkModeEnabled)
            {
                this.BackColor = Color.FromArgb(32, 32, 32);
                this.ForeColor = Color.White;
                defaultButton.BackColor = Color.FromArgb(55, 55, 55);
                defaultButton.ForeColor = Color.White;
                defaultButton.FlatStyle = FlatStyle.Flat;
                checkButton.BackColor = Color.FromArgb(55, 55, 55);
                checkButton.ForeColor = Color.White;
                checkButton.FlatStyle = FlatStyle.Flat;
            }

            this.Controls.Add(messageLabel);
            this.Controls.Add(defaultButton);
            this.Controls.Add(checkButton);
            this.AcceptButton = defaultButton;
            this.CancelButton = defaultButton;
        }
    }

    // 設定の再読み込みが発生したことを示すための内部シグナル用例外。
    // 「設定を確認」でエディタを閉じた後に最新状態で処理をやり直した場合、
    // 呼び出し元に残っている古い状態の処理を打ち切るために使う。
    private class ConfigReloadRequestedException : Exception { }

    public static void InitializeAppConfig()
    {
        try
        {
            InitializeAppConfigCore();
            // ログの設定（出力の有無、レベル、サイズ）を反映する
            AppLog.ReloadSettings();
        }
        catch (ConfigReloadRequestedException)
        {
            // 内部でInitializeAppConfig()を再実行済みのため、
            // ここでは何もせず古い状態の処理を終了する。
        }
    }

    private static void InitializeAppConfigCore()
    {
        List[PassArray] = null;
        List[ImeOffArray] = null;
        List[ImeOffTitleArray] = null;
        List[OnActivatedAppArray] = null;
        List[BackgroundArray] = null;
        val[ConfigPunctuation] = [0, 0];
        val[ConfigSpaceWidth] = [0, 0];
        keyPath[ConfigPunctuation] = @"Software\Microsoft\IME\15.0\IMEJP\MSIME";
        valueName[ConfigPunctuation] = "option1";
        valueType[ConfigPunctuation] = RegistryValueKind.DWord;
        keyPath[ConfigSpaceWidth] = @"Software\Microsoft\IME\15.0\IMEJP\MSIME";
        valueName[ConfigSpaceWidth] = "InputSpace";
        valueType[ConfigSpaceWidth] = RegistryValueKind.DWord;
        string buff = ConfigurationManager.AppSettings["AlwaysIMEMode"];
        if (!string.IsNullOrEmpty(buff))
        {
            // 互換設定
            if (buff.ToLower().CompareTo("off") == 0)
            {
                ImeModeGlobal = IMEMODE_GLOBAL;
            }
            else if (buff.ToLower().CompareTo("on") == 0)

            {
                ImeModeGlobal = IMEMODE_ON;
            }
            else if (int.TryParse(buff, out int parsedImeMode) && parsedImeMode >= 1 && parsedImeMode <= 3)
            {
                ImeModeGlobal = parsedImeMode;
            }
            else
            {
                HandleInvalidConfigValue("AlwaysIMEMode");
                ImeModeGlobal = IMEMODE_ON;
            }
        }
        else
        {
            ImeModeGlobal = IMEMODE_ON;
        }
        buff = ConfigurationManager.AppSettings["IsDarkMode"];
        if (!string.IsNullOrEmpty(buff))
        {
            if (buff.ToLower().CompareTo("on") == 0)
            {
                darkModeEnabled = true;
            }
            else
            {
                darkModeEnabled = false;
            }
        }
        else
        {
            darkModeEnabled = false;
        }
        buff = ConfigurationManager.AppSettings["PassList"];
        if (!string.IsNullOrEmpty(buff))
        {
            List[PassArray] = buff.Split(separator, StringSplitOptions.RemoveEmptyEntries);
        }
        buff = ConfigurationManager.AppSettings["ImeOffList"];
        if (!string.IsNullOrEmpty(buff))
        {
            List[ImeOffArray] = buff.Split(separator, StringSplitOptions.RemoveEmptyEntries);
        }
        buff = ConfigurationManager.AppSettings["ImeOffTitle"];
        if (!string.IsNullOrEmpty(buff))
        {
            List[ImeOffTitleArray] = buff.Split(separator, StringSplitOptions.RemoveEmptyEntries);
        }
        buff = ConfigurationManager.AppSettings["UpdateTag"];
        if (!string.IsNullOrEmpty(buff))
        {
            List[RemoveUpdateTagArray] = buff.Split(separator, StringSplitOptions.RemoveEmptyEntries);
        }
        buff = ConfigurationManager.AppSettings["intervalTime"];
        if (!string.IsNullOrEmpty(buff))
        {
            if (int.TryParse(buff, out int parsedIntervalTime))
            {
                imeInterval = parsedIntervalTime;
            }
            else
            {
                HandleInvalidConfigValue("intervalTime");
            }
        }
        buff = ConfigurationManager.AppSettings["SuspendFewTime"];
        if (!string.IsNullOrEmpty(buff))
        {
            if (int.TryParse(buff, out int parsedSuspendFewTime))
            {
                SuspendFewInterval = parsedSuspendFewTime;
            }
            else
            {
                HandleInvalidConfigValue("SuspendFewTime");
            }
        }
        buff = ConfigurationManager.AppSettings["SuspendTime"];
        if (!string.IsNullOrEmpty(buff))
        {
            if (int.TryParse(buff, out int parsedSuspendTime))
            {
                SuspendInterval = parsedSuspendTime;
            }
            else
            {
                HandleInvalidConfigValue("SuspendTime");
            }
        }
        buff = ConfigurationManager.AppSettings["NoKeyInputTime"];
        if (!string.IsNullOrEmpty(buff))
        {
            if (float.TryParse(buff, out float parsedNoKeyInputTime))
            {
                noKeyInputInterval = (int)(parsedNoKeyInputTime * 60 * 1000);
            }
            else
            {
                HandleInvalidConfigValue("NoKeyInputTime");
            }
        }
        buff = ConfigurationManager.AppSettings["DelayBackgroundTime"];
        if (!string.IsNullOrEmpty(buff))
        {
            if (float.TryParse(buff, out float parsedDelayBackgroundTime))
            {
                DelayBackgroundInterval = (int)(parsedDelayBackgroundTime * 60 * 1000 / imeInterval);
            }
            else
            {
                HandleInvalidConfigValue("DelayBackgroundTime");
            }
        }
        buff = ConfigurationManager.AppSettings["OnActivatedAppList"];
        if (!string.IsNullOrEmpty(buff))
        {
            List[OnActivatedAppArray] = buff.Split(separator, StringSplitOptions.RemoveEmptyEntries);
        }
        buff = (ConfigurationManager.AppSettings["OnActivatedAppPath"]);
        if (!string.IsNullOrEmpty(buff))
        {
            if (!File.Exists(buff))
            {
                MessageBox.Show("OnActivatedAppPath に指定したアプリが見つかりません");
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
            List[BackgroundArray] = buff.Split(separator, StringSplitOptions.RemoveEmptyEntries);
        }
        buff = (ConfigurationManager.AppSettings["BackgroundAppPath"]);
        if (!string.IsNullOrEmpty(buff))
        {
            if (!File.Exists(buff))
            {
                MessageBox.Show("BackgroundAppPath に指定したアプリが見つかりません");
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
        // MS-IME入力設定(スペース/句読点切替)の有効/無効
        // 設定画面で明示的に無効("0")にした場合のみ、レジストリへの書き込みと切替を行わない。
        // 未設定の場合は従来どおり Punctuation / SpaceWidth の設定値で動作する。
        string msImeSwitchBuff = ConfigurationManager.AppSettings["MsImeSwitch"];
        if (!string.IsNullOrEmpty(msImeSwitchBuff) && msImeSwitchBuff != "0" && msImeSwitchBuff != "1")
        {
            HandleInvalidConfigValue("MsImeSwitch");
            msImeSwitchBuff = "0";
        }
        bool msImeSwitchActive = msImeSwitchBuff != "0";
        buff = ConfigurationManager.AppSettings["Punctuation"];
        if (!string.IsNullOrEmpty(buff) && msImeSwitchActive)
        {
            int temp = 0x7FFCFFFF & (int)ReadRegistryValue(RegistryHive.CurrentUser, keyPath[ConfigPunctuation], valueName[ConfigPunctuation], valueType[ConfigPunctuation]);
            string[] parts = buff.Split(',');
            val[ConfigPunctuation] = new int[2];
            for (int i = 0; i < parts.Length; i++)
            {
                if (int.TryParse(parts[i], out int parsedPunctuation) && parsedPunctuation >= 0 && parsedPunctuation <= 3)
                {
                    val[ConfigPunctuation][i] = temp | parsedPunctuation << 16;
                }
                else
                {
                    HandleInvalidConfigValue("Punctuation");
                    val[ConfigPunctuation] = [0, 0];
                }
            }
            SetPunctuationMode = val[ConfigPunctuation][0];
            if (val[ConfigPunctuation][0] == 0 && val[ConfigPunctuation][1] == 0)
            {
                /* Nothing to do */
            }
            else
            {
                if (WriteRegistryValue(RegistryHive.CurrentUser, keyPath[ConfigPunctuation], valueName[ConfigPunctuation], SetPunctuationMode, valueType[ConfigPunctuation]))
                {
                    Debug.WriteLine($"句読点の初期化をしました");
                }
                else
                {
                    AppLog.Warn("Failed to write registory.");
                }
            }
        }
        buff = ConfigurationManager.AppSettings["SpaceWidth"];
        if (!string.IsNullOrEmpty(buff) && msImeSwitchActive)
        {
            string[] parts = buff.Split(',');
            val[ConfigSpaceWidth] = new int[2];
            for (int i = 0; i < parts.Length; i++)
            {
                if (int.TryParse(parts[i], out int parsedSpaceWidth) && parsedSpaceWidth >= 0 && parsedSpaceWidth <= 2)
                {
                    val[ConfigSpaceWidth][i] = parsedSpaceWidth;
                }
                else
                {
                    HandleInvalidConfigValue("SpaceWidth");
                    val[ConfigSpaceWidth] = [0, 0];
                }
            }
            SetSpaceMode = val[ConfigSpaceWidth][0];
            if (val[ConfigSpaceWidth][0] == 0 && val[ConfigSpaceWidth][1] == 0)
            {
                /* Nothing to do */
            }
            else
            {
                if (WriteRegistryValue(RegistryHive.CurrentUser, keyPath[ConfigSpaceWidth], valueName[ConfigSpaceWidth], SetSpaceMode, valueType[ConfigSpaceWidth]))
                {
                    Debug.WriteLine($"スペースの初期化をしました");
                }
                else
                {
                    AppLog.Warn("Failed to write registory.");
                }
            }
        }
        buff = ConfigurationManager.AppSettings["KeyAssignment"];
        if (!string.IsNullOrEmpty(buff))
        {
            if (int.TryParse(buff, out int parsedKeyAssignment) && (parsedKeyAssignment == 0 || parsedKeyAssignment == 1))
            {
                ApplyKeyAssignment(parsedKeyAssignment == 1);
            }
            else
            {
                HandleInvalidConfigValue("KeyAssignment");
            }
        }
        else
        {
            // 管理対象外になった場合は元の設定に戻す
            RestoreKeyAssignment();
        }
        buff = ConfigurationManager.AppSettings["AppExitEnd"];
        if (!string.IsNullOrEmpty(buff))
        {
            if (int.TryParse(buff, out int parsedAppExitEnd))
            {
                AppExitEndEnabled = parsedAppExitEnd == 1;
            }
            else
            {
                HandleInvalidConfigValue("AppExitEnd");
                AppExitEndEnabled = true;
            }
        }
        else
        {
            AppExitEndEnabled = true;
        }
    }
    // ============================================================
    // MS-IME入力設定
    // ============================================================

    // "初期値,切替先" 形式の設定値が、未設定("0,0"/空)以外になっているか
    internal static bool IsMsImePairConfigured(string buff)
    {
        if (string.IsNullOrWhiteSpace(buff))
            return false;
        return buff.Replace(" ", "") != "0,0";
    }

    // "初期値,切替先" 形式の設定値を解析する(各値は 0～max)
    internal static bool TryParseConfigPair(string buff, int max, out int init, out int target)
    {
        init = 0;
        target = 0;
        if (string.IsNullOrEmpty(buff))
            return false;
        string[] parts = buff.Split(',');
        return parts.Length == 2
            && int.TryParse(parts[0], out init) && init >= 0 && init <= max
            && int.TryParse(parts[1], out target) && target >= 0 && target <= max;
    }

    // MS-IMEのレジストリ値(DWORD)を読み取る。読み取れない場合は null
    internal static int? ReadMsImeDword(string name)
    {
        try
        {
            object value = ReadRegistryValue(RegistryHive.CurrentUser, MsImeRegistryPath, name, RegistryValueKind.DWord);
            return value is int intValue ? intValue : null;
        }
        catch (Exception ex)
        {
            AppLog.Warn($"ReadMsImeDword({name}) failed", ex);
            return null;
        }
    }

    // 変換キー→IME-ON / 無変換キー→IME-OFF の割り当てを有効化・無効化する
    private static void ApplyKeyAssignment(bool enable)
    {
        if (!keyAssignmentRecorded)
        {
            origKeyAssignment = ReadMsImeDword("IsKeyAssignmentEnabled") ?? -1;
            keyAssignmentRecorded = true;
        }
        bool ok;
        if (enable)
        {
            ok = WriteRegistryValue(RegistryHive.CurrentUser, MsImeRegistryPath, "IsKeyAssignmentEnabled", 1, RegistryValueKind.DWord);
            ok &= WriteRegistryValue(RegistryHive.CurrentUser, MsImeRegistryPath, "KeyAssignmentHenkan", 0, RegistryValueKind.DWord);   // 変換キー → IME-ON
            ok &= WriteRegistryValue(RegistryHive.CurrentUser, MsImeRegistryPath, "KeyAssignmentMuhenkan", 1, RegistryValueKind.DWord); // 無変換キー → IME-OFF
        }
        else
        {
            ok = WriteRegistryValue(RegistryHive.CurrentUser, MsImeRegistryPath, "IsKeyAssignmentEnabled", 0, RegistryValueKind.DWord);
        }
        if (ok)
        {
            Debug.WriteLine($"キー割り当てを{(enable ? "有効" : "無効")}にしました");
        }
        else
        {
            AppLog.Warn("Failed to write registory.");
        }
    }

    // キー割り当てを変更前の値に戻す
    private static void RestoreKeyAssignment()
    {
        if (!keyAssignmentRecorded)
            return;
        WriteRegistryValue(RegistryHive.CurrentUser, MsImeRegistryPath, "IsKeyAssignmentEnabled", origKeyAssignment >= 0 ? origKeyAssignment : 0, RegistryValueKind.DWord);
        keyAssignmentRecorded = false;
    }

    // スペース/句読点を初期値に戻す(切替が有効な場合のみ)
    // 設定の再読み込み前に呼ぶことで、無効化・値変更時にレジストリが切替先のまま残らないようにする
    private static void RestoreMsImeSwitchValues()
    {
        if (val[ConfigPunctuation] != null && val[ConfigPunctuation][0] != val[ConfigPunctuation][1])
        {
            WriteRegistryValue(RegistryHive.CurrentUser, keyPath[ConfigPunctuation], valueName[ConfigPunctuation], val[ConfigPunctuation][0], valueType[ConfigPunctuation]);
        }
        if (val[ConfigSpaceWidth] != null && val[ConfigSpaceWidth][0] != val[ConfigSpaceWidth][1])
        {
            WriteRegistryValue(RegistryHive.CurrentUser, keyPath[ConfigSpaceWidth], valueName[ConfigSpaceWidth], val[ConfigSpaceWidth][0], valueType[ConfigSpaceWidth]);
        }
    }

    // 保存したMS-IME入力設定をレジストリとトレイメニューへ反映する(設定画面の「適用」用)
    internal static void ApplyMsImeSettings()
    {
        RestoreMsImeSwitchValues();
        InitializeAppConfig();
        currentInstance?.UpdateMsImeMenu();
    }

    // トレイメニューの句読点/スペース切替項目の表示・文言を現在の設定に合わせる
    private void UpdateMsImeMenu()
    {
        bool punctSwitch = val[ConfigPunctuation][0] != val[ConfigPunctuation][1];
        bool spaceSwitch = val[ConfigSpaceWidth][0] != val[ConfigSpaceWidth][1];
        // クリックしたときに適用される値を表示する
        int nextPunct = (SetPunctuationMode == val[ConfigPunctuation][0]) ? val[ConfigPunctuation][1] : val[ConfigPunctuation][0];
        int nextSpace = (SetSpaceMode == val[ConfigSpaceWidth][0]) ? val[ConfigSpaceWidth][1] : val[ConfigSpaceWidth][0];
        menuPunctuation.Text = "「" + PunctuationText[(nextPunct >> 16) & 0x3] + "」に切替(&P)";
        menuSpace.Text = SpaceWidthText[nextSpace] + "スペースに切替(&S)";
        menuPunctuation.Visible = punctSwitch;
        separatorPunct.Visible = punctSwitch;
        menuSpace.Visible = spaceSwitch;
        separatorSpace.Visible = spaceSwitch;
    }

    private void Close_Click(object sender, EventArgs e)
    {
        DialogResult result = MessageBox.Show("AlwaysIME を終了します。", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

        if (result == DialogResult.Yes)
        {
            if (ScheduleRunBackgroundApp)
            {
                delayRunBackgroundApp = 1;
                RunBackgroundApp();
            }
            if (val[ConfigPunctuation][0] != val[ConfigPunctuation][1])
            {
                SetPunctuationMode = val[ConfigPunctuation][1];
                Punctuation_Click(this, e);
            }
            if (val[ConfigSpaceWidth][0] != val[ConfigSpaceWidth][1])
            {
                SetSpaceMode = val[ConfigSpaceWidth][1];
                Space_Click(this, e);
            }
            RestoreKeyAssignment();
            icon.Visible = false;
            icon.Dispose();
            Application.Exit();
        }
    }
    // ダークモード用のトレイメニュー描画。
    // 標準のレンダラーはセパレーターを背景色と無関係に明るい灰色で描画するため、
    // 背景を塗りつぶしてからダーク向けの色で線を描画する。
    private sealed class DarkMenuRenderer : ToolStripProfessionalRenderer
    {
        private static readonly Color MenuBackColor = Color.FromArgb(32, 32, 32);
        private static readonly Color SeparatorColor = Color.FromArgb(96, 96, 96);

        protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e)
        {
            Rectangle bounds = new Rectangle(Point.Empty, e.Item.Size);
            using (var backBrush = new SolidBrush(MenuBackColor))
            using (var linePen = new Pen(SeparatorColor))
            {
                e.Graphics.FillRectangle(backBrush, bounds);
                if (e.Vertical)
                {
                    int x = bounds.Width / 2;
                    e.Graphics.DrawLine(linePen, x, bounds.Top + 3, x, bounds.Bottom - 3);
                }
                else
                {
                    int y = bounds.Height / 2;
                    e.Graphics.DrawLine(linePen, bounds.Left + 4, y, bounds.Right - 4, y);
                }
            }
        }
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
        suspendFewMenuItem.Text = "少し無効(&D)";
        suspendFewMenuItem.Click += new EventHandler(SuspendFewMenuItem_Click);

        ToolStripMenuItem suspendMenuItem = new ToolStripMenuItem();
        suspendMenuItem.Text = "しばらく無効(&W)";
        suspendMenuItem.Click += new EventHandler(SuspendMenuItem_Click);

        ToolStripMenuItem resumeMenuItem = new ToolStripMenuItem();
        resumeMenuItem.Text = "再度有効(&R)";
        resumeMenuItem.Click += new EventHandler(ResumeMenuItem_Click);

        ToolStripSeparator separator1 = new ToolStripSeparator();

        ToolStripMenuItem updateModeMenuItem = new ToolStripMenuItem("IME Mode");
        ToolStripMenuItem menuItemModeOn = new ToolStripMenuItem("IMEオン");
        menuItemModeOn.Click += new EventHandler((sender, e) => ChangeAlwaysIMEModeAndSave(IMEMODE_ON));
        ToolStripMenuItem menuItemModeOff = new ToolStripMenuItem("IMEオフ");
        menuItemModeOff.Enabled = false;
        menuItemModeOff.Click += new EventHandler((sender, e) => ChangeAlwaysIMEModeAndSave(IMEMODE_OFF));
        ToolStripMenuItem menuItemModeGlobal = new ToolStripMenuItem("グローバル");
        menuItemModeGlobal.Click += new EventHandler((sender, e) => ChangeAlwaysIMEModeAndSave(IMEMODE_GLOBAL));

        ToolStripSeparator separator2 = new ToolStripSeparator();

        ToolStripMenuItem updateTimeMenuItem = new ToolStripMenuItem("更新間隔");
        ToolStripMenuItem menuInterval1 = new ToolStripMenuItem("100 ms");
        menuInterval1.Click += new EventHandler((sender, e) => ChangeIntervalAndSave(100));
        ToolStripMenuItem menuInterval2 = new ToolStripMenuItem("200 ms");
        menuInterval2.Click += new EventHandler((sender, e) => ChangeIntervalAndSave(200));
        ToolStripMenuItem menuInterval3 = new ToolStripMenuItem("500 ms");
        menuInterval3.Click += new EventHandler((sender, e) => ChangeIntervalAndSave(500));

        ToolStripSeparator separator3 = new ToolStripSeparator();

        ToolStripMenuItem MenuItemShowConfig = new ToolStripMenuItem();
        MenuItemShowConfig.Text = "設定を表示(&C)";
        MenuItemShowConfig.Click += MenuItemShowConfig_Click;

        ToolStripMenuItem MenuItemRegistrationDialog = new ToolStripMenuItem();
        MenuItemRegistrationDialog.Text = "IMEオフに登録";
        MenuItemRegistrationDialog.Click += MenuItemRegistrationDialog_Click;

        ToolStripMenuItem MenuItemReload = new ToolStripMenuItem();
        MenuItemReload.Text = "設定の再読み込み";
        MenuItemReload.Click += MenuReload_Click;

        ToolStripSeparator separator4 = new ToolStripSeparator();
        separatorPunct = separator4;

        menuPunctuation = new ToolStripMenuItem();
        // menuPunctuation.Text = "句読点切替(&P)";
        menuPunctuation.Text = "「" + PunctuationText[(val[ConfigPunctuation][1] >> 16) & 0x3] + "」に切替(&P)";
        menuPunctuation.Click += new EventHandler(Punctuation_Click);

        ToolStripSeparator separator5 = new ToolStripSeparator();
        separatorSpace = separator5;

        menuSpace = new ToolStripMenuItem();
        // menuSpace.Text = "スペース切替(&S)";
        menuSpace.Text = SpaceWidthText[val[ConfigSpaceWidth][1]] + "スペースに切替(&S)";

        menuSpace.Click += new EventHandler(Space_Click);

        ToolStripSeparator separator6 = new ToolStripSeparator();

        ToolStripMenuItem menuItem = new ToolStripMenuItem();
        menuItem.Text = "常駐の終了(&X)";
        menuItem.Click += new EventHandler(Close_Click);

        if (darkModeEnabled)
        {
            menu.Renderer = new DarkMenuRenderer();
            menuItemModeOn.BackColor = Color.FromArgb(32, 32, 32);
            menuItemModeOn.ForeColor = Color.White;
            menuItemModeOff.BackColor = Color.FromArgb(32, 32, 32);
            menuItemModeOff.ForeColor = Color.White;
            menuItemModeGlobal.BackColor = Color.FromArgb(32, 32, 32);
            menuItemModeGlobal.ForeColor = Color.White;
            menuInterval1.BackColor = Color.FromArgb(32, 32, 32);
            menuInterval1.ForeColor = Color.White;
            menuInterval2.BackColor = Color.FromArgb(32, 32, 32);
            menuInterval2.ForeColor = Color.White;
            menuInterval3.BackColor = Color.FromArgb(32, 32, 32);
            menuInterval3.ForeColor = Color.White;
            suspendFewMenuItem.BackColor = Color.FromArgb(32, 32, 32);
            suspendFewMenuItem.ForeColor = Color.White;
            suspendMenuItem.BackColor = Color.FromArgb(32, 32, 32);
            suspendMenuItem.ForeColor = Color.White;
            resumeMenuItem.BackColor = Color.FromArgb(32, 32, 32);
            resumeMenuItem.ForeColor = Color.White;
            updateModeMenuItem.BackColor = Color.FromArgb(32, 32, 32);
            updateModeMenuItem.ForeColor = Color.White;
            updateTimeMenuItem.BackColor = Color.FromArgb(32, 32, 32);
            updateTimeMenuItem.ForeColor = Color.White;
            MenuItemShowConfig.BackColor = Color.FromArgb(32, 32, 32);
            MenuItemShowConfig.ForeColor = Color.White;
            MenuItemRegistrationDialog.BackColor = Color.FromArgb(32, 32, 32);
            MenuItemRegistrationDialog.ForeColor = Color.White;
            MenuItemReload.BackColor = Color.FromArgb(32, 32, 32);
            MenuItemReload.ForeColor = Color.White;
            menuPunctuation.BackColor = Color.FromArgb(32, 32, 32);
            menuPunctuation.ForeColor = Color.White;
            menuSpace.BackColor = Color.FromArgb(32, 32, 32);
            menuSpace.ForeColor = Color.White;
            menuItem.BackColor = Color.FromArgb(32, 32, 32);
            menuItem.ForeColor = Color.White;
        }
        updateModeMenuItem.DropDownItems.Add(menuItemModeOn);
        updateModeMenuItem.DropDownItems.Add(menuItemModeOff);
        updateModeMenuItem.DropDownItems.Add(menuItemModeGlobal);
        updateTimeMenuItem.DropDownItems.Add(menuInterval1);
        updateTimeMenuItem.DropDownItems.Add(menuInterval2);
        updateTimeMenuItem.DropDownItems.Add(menuInterval3);
        menu.Items.Add(suspendFewMenuItem);
        menu.Items.Add(suspendMenuItem);
        menu.Items.Add(resumeMenuItem);
        menu.Items.Add(separator1);
        menu.Items.Add(updateModeMenuItem);
        if (!AppExitEndEnabled)
        {
            menu.Items.Add(separator6);
            menu.Items.Add(menuItem);
        }
        menu.Items.Add(separator2);
        menu.Items.Add(updateTimeMenuItem);
        menu.Items.Add(separator3);
        menu.Items.Add(MenuItemShowConfig);
        menu.Items.Add(MenuItemRegistrationDialog);
        menu.Items.Add(MenuItemReload);
        // 表示/非表示は UpdateMsImeMenu() で切り替える
        // (設定画面から有効/無効を変更したときに再起動なしで反映するため)
        menu.Items.Add(separator4);
        menu.Items.Add(menuPunctuation);
        menu.Items.Add(separator5);
        menu.Items.Add(menuSpace);
        if (AppExitEndEnabled)
        {
            menu.Items.Add(separator6);
            menu.Items.Add(menuItem);
        }
        UpdateMsImeMenu();
        icon.ContextMenuStrip = menu;
    }
    public class DialogForm : Form
    {
        private Label titleLabel;
        private TextBox titleTextBox;
        private Label appLabel;
        private TextBox appTextBox;
        private RadioButton titleRadioButton;
        private RadioButton appRadioButton;
        private Button okButton;
        private Button cancelButton;

        public DialogForm()
        {
            InitializeDialogComponents();
        }

        private void InitializeDialogComponents()
        {
            float UIScale;
            using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
                UIScale = graphics.DpiX / 96;
            this.Font = new Font("Meiryo", (int)(9 * Math.Pow(UIScale, 1.0 / 3.0)), FontStyle.Regular);
            this.Text = "IMEオフに登録";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new Size((int)(440 * UIScale), (int)(200 * UIScale));

            titleLabel = new Label();
            titleLabel.Text = "タイトル:";
            titleLabel.Size = new Size((int)(80 * UIScale), (int)(20 * UIScale));
            titleLabel.Location = new Point((int)(20 * UIScale), (int)(20 * UIScale));
            titleTextBox = new TextBox();
            titleTextBox.Location = new Point((int)(100 * UIScale), (int)(20 * UIScale));
            titleTextBox.Size = new Size((int)(300 * UIScale), (int)(20 * UIScale));
            titleTextBox.Text = RegistrationWindowTitle;
            appLabel = new Label();
            appLabel.Text = "アプリ名:";
            appLabel.Size = new Size((int)(80 * UIScale), (int)(20 * UIScale));
            appLabel.Location = new Point((int)(20 * UIScale), (int)(50 * UIScale));
            appTextBox = new TextBox();
            appTextBox.Location = new Point((int)(100 * UIScale), (int)(50 * UIScale));
            appTextBox.Size = new Size((int)(300 * UIScale), (int)(20 * UIScale));
            appTextBox.Text = RegistrationprocessName;
            titleRadioButton = new RadioButton();
            titleRadioButton.Text = "タイトル";
            titleRadioButton.Size = new Size((int)(90 * UIScale), (int)(30 * UIScale));
            titleRadioButton.Location = new Point((int)(100 * UIScale), (int)(80 * UIScale));
            appRadioButton = new RadioButton();
            appRadioButton.Text = "アプリ名";
            appRadioButton.Size = new Size((int)(90 * UIScale), (int)(30 * UIScale));
            appRadioButton.Location = new Point((int)(220 * UIScale), (int)(80 * UIScale));
            appRadioButton.Checked = true;
            okButton = new Button();
            okButton.Text = "登録(&R)";
            okButton.Size = new Size((int)(110 * UIScale), (int)(32 * UIScale));
            okButton.DialogResult = DialogResult.OK;
            okButton.Location = new Point((int)(100 * UIScale), (int)(110 * UIScale));
            okButton.Click += OkButton_Click;
            cancelButton = new Button();
            cancelButton.Text = "キャンセル(&C)";
            cancelButton.Size = new Size((int)(110 * UIScale), (int)(32 * UIScale));
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.Location = new System.Drawing.Point((int)(220 * UIScale), (int)(110 * UIScale));

            if (darkModeEnabled)
            {
                this.BackColor = Color.FromArgb(32, 32, 32);
                this.ForeColor = Color.White;
                titleTextBox.BackColor = Color.FromArgb(45, 45, 45);
                titleTextBox.ForeColor = Color.White;
                appTextBox.BackColor = Color.FromArgb(45, 45, 45);
                appTextBox.ForeColor = Color.White;
                okButton.BackColor = Color.FromArgb(55, 55, 55);
                okButton.ForeColor = Color.White;
                cancelButton.BackColor = Color.FromArgb(55, 55, 55);
                cancelButton.ForeColor = Color.White;
            }

            this.Controls.Add(titleLabel);
            this.Controls.Add(titleTextBox);
            this.Controls.Add(appLabel);
            this.Controls.Add(appTextBox);
            this.Controls.Add(titleRadioButton);
            this.Controls.Add(appRadioButton);
            this.Controls.Add(okButton);
            this.Controls.Add(cancelButton);
        }
        private void OkButton_Click(object sender, EventArgs e)
        {
            string selectedOption = titleRadioButton.Checked ? "Title" : "App";
            string title = titleTextBox.Text;
            string app = appTextBox.Text;

            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            if (selectedOption == "Title")
            {
                List<string> list = new List<string>(List[ImeOffTitleArray]);
                list.Add(Regex.Escape(title));
                List[ImeOffTitleArray] = list.Distinct().ToArray();
                String buff = string.Join(",", List[ImeOffTitleArray]);
                config.AppSettings.Settings["ImeOffTitle"].Value = buff;
                config.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");
                Debug.WriteLine("ImeOffTitleに追加：" + buff);
            }
            else if (selectedOption == "App")
            {
                List<string> list = new List<string>(List[ImeOffArray]);
                list.Add(app);
                List[ImeOffArray] = list.Distinct().ToArray();
                String buff = string.Join(",", List[ImeOffArray]);
                buff = string.Join(",", List[ImeOffArray]);
                config.AppSettings.Settings["ImeOffList"].Value = buff;
                config.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection("appSettings");
                Debug.WriteLine("ImeOffListに追加：" + buff);
            }
        }
    }
    private void MenuItemRegistrationDialog_Click(object sender, EventArgs e)
    {
        using (var dialog = new DialogForm())
        {
            dialog.ShowDialog();
        }
    }
    private void MenuItemShowConfig_Click(object sender, EventArgs e)
    {
        using (var configForm = new ConfigForm())
        {
            if (configForm.ShowDialog() == DialogResult.OK)
            {
                // 設定を保存した場合は再読み込みして即時反映する
                ConfigurationManager.RefreshSection("appSettings");
                InitializeAppConfig();
                UpdateMsImeMenu();
                this.timer.Interval = imeInterval;
            }
        }
    }

    private void MenuReload_Click(object sender, EventArgs e)
    {
        string configFilePath = ConfigStore.FilePath;
        try
        {
            // app.configを既定のアプリケーションで開く
            Process editorProcess = Process.Start(new ProcessStartInfo
            {
                FileName = configFilePath,
                UseShellExecute = true
            });

            if (editorProcess != null)
            {
                // プロセスの終了を待機する
                editorProcess.WaitForExit();

                // appSettingsセクションを再読み込み
                ConfigurationManager.RefreshSection("appSettings");
                RestoreMsImeSwitchValues();
                InitializeAppConfig();
                UpdateMsImeMenu();
                this.timer.Interval = imeInterval;
                // メニュー - 終了の位置も再起動後に有効
                MessageBox.Show("設定を反映しました。\n※ダークモードは再起動後に有効。", "AlwaysIME", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("設定の再読み込みに失敗しました: " + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
    private void Punctuation_Click(object sender, EventArgs e)
    {
        int newValue;
        if (SetPunctuationMode == val[ConfigPunctuation][0])
        {
            newValue = val[ConfigPunctuation][1];
            menuPunctuation.Text = "「" + PunctuationText[(val[ConfigPunctuation][0] >> 16) & 0x3] + "」に切替(&P)";
        }
        else
        {
            newValue = val[ConfigPunctuation][0];
            menuPunctuation.Text = "「" + PunctuationText[(val[ConfigPunctuation][1] >> 16) & 0x3] + "」に切替(&P)";
        }
        if (WriteRegistryValue(RegistryHive.CurrentUser, keyPath[ConfigPunctuation], valueName[ConfigPunctuation], newValue, valueType[ConfigPunctuation]))
        {
            switch ((newValue >> 16) & 0x3)
            {
                case IME_COMMA_PERIOD:
                    Debug.WriteLine($"句読点を「" + PunctuationText[IME_COMMA_PERIOD] + "」にしました 0x{0:X8}", newValue);
                    break;
                case IME_TOUTEN_KUTENN:
                    Debug.WriteLine($"句読点を「" + PunctuationText[IME_TOUTEN_KUTENN] + "」にしました 0x{0:X8}", newValue);
                    break;
                case IME_TOUTEN_PERIOD:
                    Debug.WriteLine($"句読点を「" + PunctuationText[IME_TOUTEN_PERIOD] + "」にしました 0x{0:X8}", newValue);
                    break;
                case IME_COMMA_KUTENN:
                    Debug.WriteLine($"句読点を「" + PunctuationText[IME_COMMA_KUTENN] + "」にしました 0x{0:X8}", newValue);
                    break;
                default:
                    /* Nothing to do */
                    break;
            }
            SetPunctuationMode = newValue;
        }
        else
        {
            AppLog.Warn("Failed to write registory.");
        }
    }
    private void Space_Click(object sender, EventArgs e)
    {
        int newValue;
        if (SetSpaceMode == val[ConfigSpaceWidth][0])
        {
            newValue = val[ConfigSpaceWidth][1];
            menuSpace.Text = SpaceWidthText[val[ConfigSpaceWidth][0]] + "スペースに切替(&S)";
        }
        else
        {
            newValue = val[ConfigSpaceWidth][0];
            menuSpace.Text = SpaceWidthText[val[ConfigSpaceWidth][1]] + "スペースに切替(&S)";
        }
        if (WriteRegistryValue(RegistryHive.CurrentUser, keyPath[ConfigSpaceWidth], valueName[ConfigSpaceWidth], newValue, valueType[ConfigSpaceWidth]))
        {
            switch (newValue)
            {
                case IME_AUTO_WIDTH_SPACE:
                    Debug.WriteLine($"スペースを現在の入力モードにしました");
                    break;
                case IME_FULL_WIDTH_SPACE:
                    Debug.WriteLine($"スペースを常に" + SpaceWidthText[IME_FULL_WIDTH_SPACE] + "にしました");
                    break;
                case IME_HALF_WIDTH_SPACE:
                    Debug.WriteLine($"スペースを常に" + SpaceWidthText[IME_HALF_WIDTH_SPACE] + "にしました");
                    break;
                default:
                    /* Nothing to do */
                    break;
            }
            SetSpaceMode = newValue;
        }
        else
        {
            AppLog.Warn("Failed to write registory.");
        }
    }
    private void ResumeMenuItem_Click(object sender, EventArgs e)
    {
        this.timer.Start();
        SetIcon(ICON_GREEN);
        AppLog.Info("クリックにより再開しました");
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
        AppLog.Info($"{duration.TotalMinutes}分間無効にします({DateTime.Now + duration}まで)");

        this.timer.Stop();
        await Task.Delay(duration);
        this.timer.Start();

        SetIcon(ICON_GREEN);
        AppLog.Info("再開しました");
    }
    private void ChangeAlwaysIMEModeAndSave(int mode)
    {
        // App.config の更新
        Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        config.AppSettings.Settings["AlwaysIMEMode"].Value = mode.ToString();
        config.Save(ConfigurationSaveMode.Modified);
        ConfigurationManager.RefreshSection("appSettings");

        ImeModeGlobal = mode;
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
        RunBackgroundApp();
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
                    if (foregroundprocessName.Equals(List[param][i], StringComparison.CurrentCultureIgnoreCase))
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
                    if (previousprocessName.Equals(List[param][i], StringComparison.CurrentCultureIgnoreCase))
                    {
                        return true;
                    }
                }
            }
        }
        return false;
    }
    private bool CheckRemoveUpdateTagRegex(int param)
    {
        string RegexforegroundWindowTitle = foregroundWindowTitle;
        if (List[param] != null)
        {
            if (!string.IsNullOrEmpty(foregroundWindowTitle))
            {
                for (int i = 0; i < List[param].Length; i++)
                {
                    RegexforegroundWindowTitle = Regex.Replace(RegexforegroundWindowTitle, List[RemoveUpdateTagArray][i], "").Trim();
                }
                if (RegexforegroundWindowTitle != foregroundWindowTitle)
                {
                    Debug.WriteLine($"foregroundWindowTitle:{foregroundWindowTitle}");
                    Debug.WriteLine($"previousWindowTitle--:{previousWindowTitle}");
                    Debug.WriteLine($"RegexWindowTitle-----:{RegexforegroundWindowTitle}");
                    Debug.WriteLine($"CheckRemoveUpdateTagRegex:{(previousWindowTitle == RegexforegroundWindowTitle ? "成功" : "失敗")}");
                    return true;
                }
            }
        }
        return false;
    }
    private void RemoveUpdateTagRegex()
    {
        if (List[RemoveUpdateTagArray] != null)
        {
            if (!string.IsNullOrEmpty(foregroundWindowTitle))
            {
                for (int i = 0; i < List[RemoveUpdateTagArray].Length; i++)
                {
                    foregroundWindowTitle = Regex.Replace(foregroundWindowTitle, List[RemoveUpdateTagArray][i], "").Trim();
                }
            }
        }
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
                Debug.WriteLine($"{noKeyInputInterval / 1000}秒間キーボードやマウス入力がありません");
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
                AppLog.Warn("IMEが有効になりません");
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
                AppLog.Warn("IMEが無効になりません");
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
            AppLog.Warn("SetImeOffList：IMEが無効になりません");
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
            AppLog.Warn("IMEが有効になりません");
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
                    AppLog.Info("連携アプリを起動します: " + Path.GetFileName(RunOnActivatedAppPath));
                    Process.Start(RunOnActivatedAppPath, RunOnActivatedArgv);
                    flagOnActivated = true;
                    SetIcon(ICON_RED);
                }
            }
        }
    }
    private void MonitorBackground(int param)
    {
        // 利用状況が分かる情報のため、ファイルログには出さない（デバッガー専用）
        Debug.WriteLine($"{previousprocessName},{foregroundprocessName}");
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
    static object ReadRegistryValue(RegistryHive hive, string keyPath, string valueName, RegistryValueKind valueType)
    {
        using (var regKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default).OpenSubKey(keyPath))
        {
            if (regKey != null)
            {
                object value = regKey.GetValue(valueName, null);
                if (value != null && regKey.GetValueKind(valueName) == valueType)
                {
                    return value;
                }
            }
        }
        return null;
    }
    static bool WriteRegistryValue(RegistryHive hive, string keyPath, string valueName, object value, RegistryValueKind valueType)
    {
        try
        {
            using (var regKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default).CreateSubKey(keyPath))
            {
                if (regKey != null)
                {
                    regKey.SetValue(valueName, value, valueType);
                    return true;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error writing registry value: {ex.Message}");
        }
        return false;
    }
    // 半角カナ/全角英数/カタカナ モードを強制的に「ひらがな」モードに変更する
    void MonitorActiveWindow()
    {
        // IME状態の取得
        GUITHREADINFO gti = new GUITHREADINFO();
        gti.cbSize = Marshal.SizeOf(gti);

        if (!GetGUIThreadInfo(0, ref gti))
        {
            AppLog.Warn("GetGUIThreadInfo failed");
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
            AppLog.Info("IMEが無効です");
            changeIme = false;
            SetIcon(ICON_GREEN);
            return;
        }

        // キーボードやマウス未入力ならIMEの状態を復元する
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
                            if (CheckRemoveUpdateTagRegex(RemoveUpdateTagArray))
                            {
                                RemoveUpdateTagRegex();
                                Debug.WriteLine($"タイトルから更新タグを削除しました。");
                            }
                            Debug.WriteLine($"タイトル:{foregroundWindowTitle} プロセス名:{foregroundprocessName}");
                        }
                        else
                        {
                            foregroundWindowTitle = "";
                            // ウィンドウタイトルやプロセス名はファイルログには出さない（デバッガー専用）
                            Debug.WriteLine($"タイトルを取得できません。タイトル:{foregroundWindowTitle} プロセス名:{foregroundprocessName}");
                            /* タイトルが空欄のプロセスが存在するので後でreturn;をする */
                            // return;
                        }
                    }
                    catch (Exception ex)
                    {
                        AppLog.Warn("GetProcessById Failed", ex);
                        return;
                    }
                }
                else
                {
                    AppLog.Warn("GetWindowThreadProcessId Failed");
                    return;
                }
            }
            catch (Exception ex)
            {
                AppLog.Warn("Getting window information Failed", ex);
                return;
            }
        }
        else
        {
            AppLog.Warn("GetForegroundWindow Failed");
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
        if (string.IsNullOrEmpty(foregroundWindowTitle))
        {
            /* 先に出力してある */
            // Trace.WriteLine($"タイトルを取得できません。タイトル:{foregroundWindowTitle} プロセス名:{foregroundprocessName}");
            SetIcon(ICON_GREEN);
            return;
        }
        if (foregroundWindowTitle != previousWindowTitle)
        {
            if (ImeModeGlobal == IMEMODE_GLOBAL)
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
            if (ImeModeGlobal == IMEMODE_ON)
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
            if (ImeModeGlobal == IMEMODE_OFF)
            {
                if (CheckforegroundprocessName(ImeOffArray))
                {
                    SetImePreset();
                }
                else if (CheckforegroundWindowTitleRegex(ImeOffTitleArray))
                {
                    SetImePreset();
                }
                else
                {
                    SetImeOffList();
                }
            }
            // タスクバーをクリックするとタイトルが空欄のexplorerがヒット
            // ＆ 自身のプロセスで上書きしない
            if (foregroundprocessName != "explorer" && foregroundprocessName != "AlwaysIME")
            {
                RegistrationWindowTitle = foregroundWindowTitle;
                RegistrationprocessName = foregroundprocessName;
            }
        }
        if (imeEnabled)
        {
            switch (imeConvMode)
            {
                case CMode_Hiragana:
                case CModeMS_Hiragana:
                    /* Nothing to do */
                    break;
                case CMode_HankakuKana: /* through */
                case CMode_ZenkakuEisu: /* through */
                case CMode_ZenkakuKana: /* through */
                case CModeMS_HankakuKana: /* through */
                case CModeMS_ZenkakuEisu: /* through */
                case CModeMS_ZenkakuKana:
                    IntPtr result = SendMessage(imwd, WM_IME_CONTROL, (IntPtr)IMC_SETCONVERSIONMODE, (IntPtr)CMode_Hiragana); // ひらがなモードに設定
                    int error = Marshal.GetLastWin32Error();
                    Console.WriteLine("SendMessage result: " + result.ToInt64());
                    Console.WriteLine("Last Win32 Error: " + error);
                    break;
                default:
                    Debug.WriteLine($"不明な status code:{imeConvMode}");
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
