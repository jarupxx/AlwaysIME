using Microsoft.Win32;
using System.Collections.Generic;
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
            ApplicationConfiguration.Initialize();
            BackupConfigFile();
            ResidentTest.InitializeAppConfig();
            ResidentTest rm = new ResidentTest();
            WatchForRestartSignal();
            Application.Run();
            restartSignalRegisteredWait?.Unregister(null);
            restartSignalEvent?.Dispose();
            mutex.ReleaseMutex();
        }
        else
        {
            // すでに起動しているインスタンスがある場合は、再起動シグナルを送信
            try
            {
                using EventWaitHandle signalEvent = EventWaitHandle.OpenExisting(RestartSignalEventName);
                signalEvent.Set();
            }
            catch (WaitHandleCannotBeOpenedException)
            {
                // 既存インスタンスがまだイベントを作成していない場合は何もできない
                Trace.WriteLine("再起動シグナル用イベントが見つかりません。");
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
        string exePath = Process.GetCurrentProcess().MainModule.FileName;
        Process.Start(exePath);

        // 現在のプロセスを終了
        Environment.Exit(0);
    }

    static void BackupConfigFile()
    {
        string exePath = AppDomain.CurrentDomain.BaseDirectory;
        string configFilePath = Path.Combine(exePath, "AlwaysIME.dll.config");
        string backupFilePath = Path.Combine(exePath, "~AlwaysIME.dll.config.bak");

        if (File.Exists(configFilePath))
        {
            DateTime configLastWriteTime = File.GetLastWriteTime(configFilePath);

            if (!File.Exists(backupFilePath))
            {
                // app.config.bakが存在しない場合、新しく作成
                File.Copy(configFilePath, backupFilePath, true);
                Debug.WriteLine("新しくバックアップを作成しました。");
            }
            else
            {
                DateTime backupLastWriteTime = File.GetLastWriteTime(backupFilePath);

                // 設定よりバックアップが新しい場合は起動しない
                if (backupLastWriteTime > configLastWriteTime)
                {
                    MessageBox.Show($"設定ファイルとバックアップを確認してください。\n\n場所: {exePath}\n\"AlwaysIME.dll.config\"\t日時: {configLastWriteTime:yyyy/MM/dd HH:mm:ss}\n\"~AlwaysIME.dll.config.bak\"\t日時: {backupLastWriteTime:yyyy/MM/dd HH:mm:ss}", "AlwaysIME", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    Environment.Exit(0);
                }
                else if ((DateTime.Now - configLastWriteTime).TotalDays >= 2)
                {
                    File.Copy(configFilePath, backupFilePath, true);
                    Debug.WriteLine("2日経過したのでバックアップを作成しました。");
                }
                else
                {
                    Debug.WriteLine("今日のバックアップは済んでいます。");
                }
            }
        }
        else
        {
            Trace.WriteLine("AlwaysIME.dll.configが見つかりません。");
        }
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
    static bool darkModeEnabled = false;
    static bool previousimeEnabled = true;
    static bool AppExitEndEnabled = true;
    static bool changeIme = false;
    static bool noKeyInput = false;
    static bool flagOnActivated = false;
    static bool ScheduleRunBackgroundApp = false;
    static int delayRunBackgroundApp = 2147483647;
    static string foregroundprocessName;
    private static string foregroundWindowTitle;
    static string RegistrationWindowTitle;
    static readonly string[][] List = new string[6][];
    const int PassArray = 0;
    const int ImeOffArray = 1;
    const int ImeOffTitleArray = 2;
    const int RemoveUpdateTagArray = 3;
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
    private static int noKeyInputInterval = 6000;
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
    const int IME_AUTO_WIDTH_SPACE = 0;
    const int IME_FULL_WIDTH_SPACE = 1;
    const int IME_HALF_WIDTH_SPACE = 2;
    // 0: 現在の入力モード
    // 1: 常に全角
    // 2: 常に半角

    private readonly string[] PunctuationText = ["，．", "、。", "、．", "，。"];
    const int IME_COMMA_PERIOD = 0;
    const int IME_TOUTEN_KUTENN = 1;
    const int IME_TOUTEN_PERIOD = 2;
    const int IME_COMMA_KUTENN = 3;
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
    private static readonly Dictionary<string, string> DefaultConfigValues = new Dictionary<string, string>
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
            var map = new ExeConfigurationFileMap { ExeConfigFilename = "AlwaysIME.dll.config" };
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
        var map = new ExeConfigurationFileMap { ExeConfigFilename = "AlwaysIME.dll.config" };
        string configFilePath = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None).FilePath;

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
                    Trace.WriteLine("Failed to write registory.");
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
                    Trace.WriteLine("Failed to write registory.");
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
    private static bool IsMsImePairConfigured(string buff)
    {
        if (string.IsNullOrWhiteSpace(buff))
            return false;
        return buff.Replace(" ", "") != "0,0";
    }

    // "初期値,切替先" 形式の設定値を解析する(各値は 0～max)
    private static bool TryParseConfigPair(string buff, int max, out int init, out int target)
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
    private static int? ReadMsImeDword(string name)
    {
        try
        {
            object value = ReadRegistryValue(RegistryHive.CurrentUser, MsImeRegistryPath, name, RegistryValueKind.DWord);
            return value is int intValue ? intValue : null;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"ReadMsImeDword({name}) failed: {ex.Message}");
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
            Trace.WriteLine("Failed to write registory.");
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
    private static void ApplyMsImeSettings()
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
            separator1.BackColor = Color.FromArgb(32, 32, 32);
            separator1.ForeColor = Color.White;
            updateModeMenuItem.BackColor = Color.FromArgb(32, 32, 32);
            updateModeMenuItem.ForeColor = Color.White;
            separator2.BackColor = Color.FromArgb(32, 32, 32);
            separator2.ForeColor = Color.White;
            updateTimeMenuItem.BackColor = Color.FromArgb(32, 32, 32);
            updateTimeMenuItem.ForeColor = Color.White;
            separator3.BackColor = Color.FromArgb(32, 32, 32);
            separator3.ForeColor = Color.White;
            MenuItemShowConfig.BackColor = Color.FromArgb(32, 32, 32);
            MenuItemShowConfig.ForeColor = Color.White;
            MenuItemRegistrationDialog.BackColor = Color.FromArgb(32, 32, 32);
            MenuItemRegistrationDialog.ForeColor = Color.White;
            MenuItemReload.BackColor = Color.FromArgb(32, 32, 32);
            MenuItemReload.ForeColor = Color.White;
            separator4.BackColor = Color.FromArgb(32, 32, 32);
            separator4.ForeColor = Color.White;
            menuPunctuation.BackColor = Color.FromArgb(32, 32, 32);
            menuPunctuation.ForeColor = Color.White;
            separator5.BackColor = Color.FromArgb(32, 32, 32);
            separator5.ForeColor = Color.White;
            menuSpace.BackColor = Color.FromArgb(32, 32, 32);
            menuSpace.ForeColor = Color.White;
            separator6.BackColor = Color.FromArgb(32, 32, 32);
            separator6.ForeColor = Color.White;
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
        if (!darkModeEnabled)
            menu.Items.Add(separator1);
        menu.Items.Add(updateModeMenuItem);
        if (!AppExitEndEnabled)
        {
            if (!darkModeEnabled)
                menu.Items.Add(separator6);
            menu.Items.Add(menuItem);
        }
        if (!darkModeEnabled)
            menu.Items.Add(separator2);
        menu.Items.Add(updateTimeMenuItem);
        if (!darkModeEnabled)
            menu.Items.Add(separator3);
        menu.Items.Add(MenuItemShowConfig);
        menu.Items.Add(MenuItemRegistrationDialog);
        menu.Items.Add(MenuItemReload);
        // 表示/非表示は UpdateMsImeMenu() で切り替える
        // (設定画面から有効/無効を変更したときに再起動なしで反映するため)
        if (!darkModeEnabled)
            menu.Items.Add(separator4);
        menu.Items.Add(menuPunctuation);
        if (!darkModeEnabled)
            menu.Items.Add(separator5);
        menu.Items.Add(menuSpace);
        if (AppExitEndEnabled)
        {
            if (!darkModeEnabled)
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

    // ============================================================
    // 設定ウィンドウ
    // カテゴリー一覧（左） + 選択カテゴリーの編集パネル（右）の
    // 2ペイン構成。リスト系設定はその場で追加・編集・削除・並べ替えができる。
    // ============================================================
    public class ConfigForm : Form
    {
        private class ConfigCategory
        {
            public string AppSettingKey;   // App.configのキー名（"seconds"/"advanced"タイプではnull可）
            public string Label;           // 左のカテゴリー一覧に表示する名前
            public string Desc;            // 説明文
            public string Type;            // "list" / "seconds" / "msime" / "advanced"
            public int ListIndex = -1;     // "list"タイプの場合、List[] 配列のインデックス
        }

        private readonly ConfigCategory[] categories = new ConfigCategory[]
        {
            new ConfigCategory
            {
                AppSettingKey = "PassList",
                Label = "制御しないアプリ",
                Desc = "IMEを一切操作しないアプリです。\r\n実行ファイル名（拡張子なし）を1行1件で入力してください。\r\n例: AlwaysIME",
                Type = "list",
                ListIndex = PassArray,
            },
            new ConfigCategory
            {
                AppSettingKey = "ImeOffList",
                Label = "IME-OFFアプリ",
                Desc = "アプリの切り替えやウィンドウタイトルの変化時に、IMEをOFFにするアプリです。\r\n実行ファイル名（拡張子なし）を1行1件で入力してください。\r\n例: cmd, powershell",
                Type = "list",
                ListIndex = ImeOffArray,
            },
            new ConfigCategory
            {
                AppSettingKey = "ImeOffTitle",
                Label = "IME-OFFタイトルパターン",
                Desc = "ウィンドウタイトルにマッチしたら、IMEをOFFにする正規表現です。\r\n1行1パターン、全アプリ共通で適用されます。\r\n例: \\.cs$",
                Type = "list",
                ListIndex = ImeOffTitleArray,
            },
            new ConfigCategory
            {
                AppSettingKey = "UpdateTag",
                Label = "タイトル除外タグ",
                Desc = "ウィンドウタイトルの変化を判定する際に、無視する部分を表す正規表現です。\r\nマッチした部分を除去してから変化を判定します。\r\n例: \\(更新\\)",
                Type = "list",
                ListIndex = RemoveUpdateTagArray,
            },
            new ConfigCategory
            {
                AppSettingKey = "NoKeyInputTime",
                Label = "未入力タイムアウト",
                Desc = "この時間、キーボードやマウスの入力が無かったら、IMEの制御状態を復元します。\r\n秒単位で入力してください。",
                Type = "seconds",
            },
            new ConfigCategory
            {
                AppSettingKey = null,
                Label = "MS-IME入力設定",
                Desc = "MS-IMEのスペース・句読点をレジストリ経由で制御します。\r\n有効にするとトレイメニューから切替できます。\r\n終了時はレジストリを元の値に復元します。",
                Type = "msime",
            },
            new ConfigCategory
            {
                AppSettingKey = null,
                Label = "上級者向け設定",
                Desc = "動作に詳しい方向けの設定です。\r\n通常は変更不要です。",
                Type = "advanced",
            },
        };

        // カテゴリーごとの作業用リスト（"list"タイプのみ使用）
        private readonly Dictionary<int, List<string>> categoryItems = new Dictionary<int, List<string>>();
        private int selectedCategoryIndex = 0;
        private int editingIndex = -1; // -1 = 新規追加モード
        private float uiScale = 1f;

        // ---- コントロール ----
        private ListBox catList;
        private Label descLabel;

        private ListBox itemList;
        private Button btnEdit, btnDelete, btnUp, btnDown;
        private Label valLabel;
        private TextBox inputEdit;
        private Button btnAddOK;

        private Label secondsLabel;
        private NumericUpDown secondsCtrl;

        private CheckBox chkAppExitEnd;
        private CheckBox chkDarkMode;
        private Label advNote;

        // MS-IME入力設定パネル
        private static readonly string[] InputSpaceLabels = ["現在の入力モード", "常に全角", "常に半角"];
        private static readonly string[] PunctuationLabels = ["「，．」", "「、。」", "「、．」", "「，。」"];
        private CheckBox chkKeyAssign;
        private CheckBox chkMsImeSwitch;
        private ComboBox cmbSpaceInit, cmbSpaceTo, cmbPunctInit, cmbPunctTo;
        private Control[] msimeControls;

        private Button btnSave, btnApply, btnCancel;

        public ConfigForm()
        {
            InitializeComponents();
            LoadCurrentValues();
            RenderCategory(0);
        }

        private void InitializeComponents()
        {
            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
                uiScale = g.DpiX / 96;

            int W = (int)(640 * uiScale);
            int H = (int)(430 * uiScale);
            int LW = (int)(170 * uiScale);
            int SP = (int)(8 * uiScale);

            this.Text = "AlwaysIME 設定";
            this.Font = new Font("Meiryo UI", 9);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new Size(W, H);

            if (darkModeEnabled)
            {
                this.BackColor = Color.FromArgb(32, 32, 32);
                this.ForeColor = Color.White;
            }
            else
            {
                this.BackColor = Color.FromArgb(240, 240, 240);
            }

            // ---- 左側：カテゴリー一覧 ----
            catList = new ListBox();
            catList.Location = new Point(SP, SP);
            catList.Size = new Size(LW - SP * 2, H - SP * 2 - (int)(40 * uiScale));
            catList.IntegralHeight = false;
            foreach (var cat in categories)
                catList.Items.Add(cat.Label);
            catList.SelectedIndexChanged += CatList_SelectedIndexChanged;
            this.Controls.Add(catList);

            int RX = LW + SP;
            int RW = W - RX - SP;

            // ---- 右側：説明文 ----
            descLabel = new Label();
            descLabel.Location = new Point(RX, SP);
            descLabel.Size = new Size(RW, (int)(56 * uiScale));
            descLabel.AutoSize = false;
            this.Controls.Add(descLabel);

            int panelTop = SP + (int)(60 * uiScale);
            int panelHeight = H - panelTop - (int)(48 * uiScale);

            // ---- リスト編集パネル ----
            int btnW = (int)(76 * uiScale);
            itemList = new ListBox();
            itemList.Location = new Point(RX, panelTop);
            itemList.Size = new Size(RW - btnW - SP, panelHeight - (int)(32 * uiScale));
            itemList.IntegralHeight = false;
            itemList.DoubleClick += (s, e) => StartEdit();
            this.Controls.Add(itemList);

            int BX = RX + RW - btnW;
            btnEdit = new Button { Text = "編集", Location = new Point(BX, panelTop), Size = new Size(btnW, (int)(24 * uiScale)) };
            btnDelete = new Button { Text = "削除", Location = new Point(BX, panelTop + (int)(28 * uiScale)), Size = new Size(btnW, (int)(24 * uiScale)) };
            btnUp = new Button { Text = "↑ 上へ", Location = new Point(BX, panelTop + (int)(60 * uiScale)), Size = new Size(btnW, (int)(24 * uiScale)) };
            btnDown = new Button { Text = "↓ 下へ", Location = new Point(BX, panelTop + (int)(88 * uiScale)), Size = new Size(btnW, (int)(24 * uiScale)) };
            btnEdit.Click += (s, e) => StartEdit();
            btnDelete.Click += (s, e) => DeleteItem();
            btnUp.Click += (s, e) => MoveItem(-1);
            btnDown.Click += (s, e) => MoveItem(1);
            this.Controls.Add(btnEdit);
            this.Controls.Add(btnDelete);
            this.Controls.Add(btnUp);
            this.Controls.Add(btnDown);

            int inputY = panelTop + panelHeight - (int)(28 * uiScale);
            valLabel = new Label { Text = "値：", Location = new Point(RX, inputY + (int)(4 * uiScale)), Size = new Size((int)(28 * uiScale), (int)(20 * uiScale)) };
            inputEdit = new TextBox { Location = new Point(RX + (int)(30 * uiScale), inputY), Size = new Size(RW - btnW - SP - (int)(34 * uiScale), (int)(24 * uiScale)) };
            btnAddOK = new Button { Text = "リストへ追加", Location = new Point(BX, inputY), Size = new Size(btnW, (int)(24 * uiScale)) };
            btnAddOK.Click += (s, e) => CommitItem();
            this.Controls.Add(valLabel);
            this.Controls.Add(inputEdit);
            this.Controls.Add(btnAddOK);

            // ---- 秒数パネル ----
            secondsLabel = new Label { Text = "秒数：", Location = new Point(RX, panelTop), Size = new Size((int)(60 * uiScale), (int)(24 * uiScale)) };
            secondsCtrl = new NumericUpDown
            {
                Location = new Point(RX + (int)(64 * uiScale), panelTop),
                Size = new Size((int)(80 * uiScale), (int)(24 * uiScale)),
                Minimum = 1,
                Maximum = 3600,
                Value = 45,
            };
            this.Controls.Add(secondsLabel);
            this.Controls.Add(secondsCtrl);

            // ---- 上級者向け設定パネル ----
            chkAppExitEnd = new CheckBox
            {
                Text = "トレイメニューに「常駐の終了」を表示する",
                Location = new Point(RX, panelTop),
                Size = new Size(RW, (int)(24 * uiScale)),
            };
            chkDarkMode = new CheckBox
            {
                Text = "トレイメニューをダークモードにする（再起動後に反映）",
                Location = new Point(RX, panelTop + (int)(28 * uiScale)),
                Size = new Size(RW, (int)(24 * uiScale)),
            };
            advNote = new Label
            {
                Text = "※ ここで変更した内容は「保存して閉じる」または「適用」を押すまで反映されません。",
                Location = new Point(RX, panelTop + (int)(60 * uiScale)),
                Size = new Size(RW, (int)(40 * uiScale)),
            };
            this.Controls.Add(chkAppExitEnd);
            this.Controls.Add(chkDarkMode);
            this.Controls.Add(advNote);

            // ---- MS-IME入力設定パネル ----
            chkKeyAssign = new CheckBox
            {
                Text = "変換キー→IME-ON / 無変換キー→IME-OFFに割り当てる",
                Location = new Point(RX, panelTop),
                Size = new Size(RW, (int)(24 * uiScale)),
            };
            chkMsImeSwitch = new CheckBox
            {
                Text = "スペース / 句読点切替を有効にする",
                Location = new Point(RX, panelTop + (int)(28 * uiScale)),
                Size = new Size(RW, (int)(24 * uiScale)),
            };

            // 左カラム：スペース設定  右カラム：句読点設定
            int colGap = (int)(16 * uiScale);
            int halfW = (RW - colGap) / 2;
            int colL = RX;
            int colR = RX + halfW + colGap;
            int lblW = (int)(56 * uiScale);
            int ddX = lblW + (int)(4 * uiScale);
            int ddW = (int)((halfW - ddX) * 0.8);
            int rowH = (int)(20 * uiScale);
            int yHeader = panelTop + (int)(62 * uiScale);
            int yInit = panelTop + (int)(81 * uiScale);
            int yTo = panelTop + (int)(109 * uiScale);
            int lblDy = (int)(4 * uiScale);

            Label lblSpaceHeader = MakeMsImeLabel("── スペース ──────────", colL, yHeader, halfW, (int)(18 * uiScale));
            Label lblSpaceInit = MakeMsImeLabel("初期値：", colL, yInit + lblDy, lblW, rowH);
            cmbSpaceInit = MakeMsImeCombo(InputSpaceLabels, colL + ddX, yInit, ddW);
            Label lblSpaceTo = MakeMsImeLabel("切替先：", colL, yTo + lblDy, lblW, rowH);
            cmbSpaceTo = MakeMsImeCombo(InputSpaceLabels, colL + ddX, yTo, ddW);

            Label lblPunctHeader = MakeMsImeLabel("── 句読点 ──────────", colR, yHeader, halfW, (int)(18 * uiScale));
            Label lblPunctInit = MakeMsImeLabel("初期値：", colR, yInit + lblDy, lblW, rowH);
            cmbPunctInit = MakeMsImeCombo(PunctuationLabels, colR + ddX, yInit, ddW);
            Label lblPunctTo = MakeMsImeLabel("切替先：", colR, yTo + lblDy, lblW, rowH);
            cmbPunctTo = MakeMsImeCombo(PunctuationLabels, colR + ddX, yTo, ddW);

            Label msimeHint = MakeMsImeLabel(
                "初期値：起動時・保存時にMS-IMEへ書き込む設定値です。\r\n" +
                "切替先：トレイメニューの切替項目を押したときの値です。もう一度押すと初期値に戻ります。\r\n" +
                "終了時はスペース・句読点を初期値に、キー割り当てを変更前の設定に戻します。",
                RX, panelTop + (int)(144 * uiScale), RW, (int)(110 * uiScale));

            msimeControls = new Control[]
            {
                chkKeyAssign, chkMsImeSwitch,
                lblSpaceHeader, lblSpaceInit, cmbSpaceInit, lblSpaceTo, cmbSpaceTo,
                lblPunctHeader, lblPunctInit, cmbPunctInit, lblPunctTo, cmbPunctTo,
                msimeHint,
            };
            this.Controls.AddRange(msimeControls);

            // ---- 下部ボタン ----
            int btnY = H - (int)(36 * uiScale);
            btnSave = new Button { Text = "保存して閉じる", Location = new Point(RX, btnY), Size = new Size((int)(120 * uiScale), (int)(28 * uiScale)) };
            btnApply = new Button { Text = "適用", Location = new Point(RX + (int)(124 * uiScale), btnY), Size = new Size((int)(80 * uiScale), (int)(28 * uiScale)) };
            btnCancel = new Button { Text = "キャンセル", Location = new Point(W - SP - (int)(80 * uiScale), btnY), Size = new Size((int)(80 * uiScale), (int)(28 * uiScale)) };
            btnSave.Click += (s, e) => { if (SaveAll()) { this.DialogResult = DialogResult.OK; this.Close(); } };
            btnApply.Click += (s, e) => SaveAll();
            btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };
            this.AcceptButton = btnSave;
            this.Controls.Add(btnSave);
            this.Controls.Add(btnApply);
            this.Controls.Add(btnCancel);

            if (darkModeEnabled)
            {
                foreach (Control c in this.Controls)
                {
                    if (c is Label || c is CheckBox)
                        c.ForeColor = Color.White;
                    if (c is Button || c is TextBox || c is ListBox || c is NumericUpDown || c is ComboBox)
                    {
                        c.BackColor = Color.FromArgb(45, 45, 45);
                        c.ForeColor = Color.White;
                    }
                }
            }
        }

        private static Label MakeMsImeLabel(string text, int x, int y, int w, int h)
        {
            return new Label { Text = text, Location = new Point(x, y), Size = new Size(w, h) };
        }

        private static ComboBox MakeMsImeCombo(string[] items, int x, int y, int w)
        {
            var cmb = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(x, y),
                Width = w,
            };
            cmb.Items.AddRange(items);
            return cmb;
        }

        private static void SetAppSetting(Configuration config, string key, string value)
        {
            if (config.AppSettings.Settings[key] != null)
                config.AppSettings.Settings[key].Value = value;
            else
                config.AppSettings.Settings.Add(key, value);
        }

        private void LoadCurrentValues()
        {
            foreach (var cat in categories)
            {
                if (cat.Type == "list")
                {
                    string[] source = List[cat.ListIndex];
                    categoryItems[Array.IndexOf(categories, cat)] = source != null
                        ? new List<string>(source)
                        : new List<string>();
                }
            }
            secondsCtrl.Value = Math.Max(secondsCtrl.Minimum, Math.Min(secondsCtrl.Maximum, noKeyInputInterval / 1000));
            chkAppExitEnd.Checked = AppExitEndEnabled;
            chkDarkMode.Checked = darkModeEnabled;

            // ---- MS-IME入力設定 ----
            // キー割り当て: 設定済みならその値、未設定ならレジストリの現在値
            string keyAssignBuff = ConfigurationManager.AppSettings["KeyAssignment"];
            chkKeyAssign.Checked = !string.IsNullOrEmpty(keyAssignBuff)
                ? keyAssignBuff == "1"
                : ReadMsImeDword("IsKeyAssignmentEnabled") == 1;

            // 切替の有効/無効: 明示設定があればその値、未設定なら従来の設定値の有無で判定
            string switchBuff = ConfigurationManager.AppSettings["MsImeSwitch"];
            string punctBuff = ConfigurationManager.AppSettings["Punctuation"];
            string spaceBuff = ConfigurationManager.AppSettings["SpaceWidth"];
            chkMsImeSwitch.Checked = string.IsNullOrEmpty(switchBuff)
                ? (IsMsImePairConfigured(punctBuff) || IsMsImePairConfigured(spaceBuff))
                : switchBuff == "1";

            // 初期値/切替先: 設定済みならその値、未設定ならレジストリの現在値を初期値にする
            int regSpace = ReadMsImeDword("InputSpace") ?? IME_AUTO_WIDTH_SPACE;
            if (regSpace < IME_AUTO_WIDTH_SPACE || regSpace > IME_HALF_WIDTH_SPACE)
                regSpace = IME_AUTO_WIDTH_SPACE;
            int? regOption1 = ReadMsImeDword("option1");
            int regPunct = regOption1.HasValue ? (regOption1.Value >> 16) & 0x3 : IME_TOUTEN_KUTENN;

            int spaceInit = regSpace;
            int spaceTo = regSpace == IME_HALF_WIDTH_SPACE ? IME_FULL_WIDTH_SPACE : IME_HALF_WIDTH_SPACE;
            if (TryParseConfigPair(spaceBuff, IME_HALF_WIDTH_SPACE, out int cfgSpaceInit, out int cfgSpaceTo)
                && !(cfgSpaceInit == 0 && cfgSpaceTo == 0))
            {
                spaceInit = cfgSpaceInit;
                spaceTo = cfgSpaceTo;
            }

            int punctInit = regPunct;
            int punctTo = IME_COMMA_KUTENN;
            if (TryParseConfigPair(punctBuff, IME_COMMA_KUTENN, out int cfgPunctInit, out int cfgPunctTo)
                && !(cfgPunctInit == 0 && cfgPunctTo == 0))
            {
                punctInit = cfgPunctInit;
                punctTo = cfgPunctTo;
            }

            cmbSpaceInit.SelectedIndex = spaceInit;
            cmbSpaceTo.SelectedIndex = spaceTo;
            cmbPunctInit.SelectedIndex = punctInit;
            cmbPunctTo.SelectedIndex = punctTo;
        }

        private void CatList_SelectedIndexChanged(object sender, EventArgs e)
        {
            int idx = catList.SelectedIndex;
            if (idx < 0)
                return;
            ConfirmPendingInput();
            selectedCategoryIndex = idx;
            RenderCategory(idx);
        }

        private void RenderCategory(int idx)
        {
            catList.SelectedIndexChanged -= CatList_SelectedIndexChanged;
            catList.SelectedIndex = idx;
            catList.SelectedIndexChanged += CatList_SelectedIndexChanged;

            var cat = categories[idx];
            descLabel.Text = cat.Desc;

            bool isList = cat.Type == "list";
            bool isSeconds = cat.Type == "seconds";
            bool isAdvanced = cat.Type == "advanced";
            bool isMsIme = cat.Type == "msime";

            itemList.Visible = isList;
            btnEdit.Visible = isList;
            btnDelete.Visible = isList;
            btnUp.Visible = isList;
            btnDown.Visible = isList;
            valLabel.Visible = isList;
            inputEdit.Visible = isList;
            btnAddOK.Visible = isList;

            secondsLabel.Visible = isSeconds;
            secondsCtrl.Visible = isSeconds;

            chkAppExitEnd.Visible = isAdvanced;
            chkDarkMode.Visible = isAdvanced;
            advNote.Visible = isAdvanced;

            foreach (Control c in msimeControls)
                c.Visible = isMsIme;

            if (isList)
            {
                editingIndex = -1;
                inputEdit.Text = "";
                btnAddOK.Text = "リストへ追加";
                RefreshListBox(0);
            }
        }

        private List<string> CurrentItems => categoryItems[selectedCategoryIndex];

        private void RefreshListBox(int selectOneBased)
        {
            var items = CurrentItems;
            itemList.BeginUpdate();
            itemList.Items.Clear();
            foreach (var v in items)
                itemList.Items.Add(v);
            itemList.EndUpdate();
            if (selectOneBased > 0 && selectOneBased <= items.Count)
                itemList.SelectedIndex = selectOneBased - 1;
            else if (items.Count > 0)
                itemList.SelectedIndex = 0;
        }

        private void StartEdit()
        {
            int idx = itemList.SelectedIndex;
            if (idx < 0)
                return;
            editingIndex = idx;
            inputEdit.Text = CurrentItems[idx];
            btnAddOK.Text = "更新";
            inputEdit.Focus();
        }

        private void CommitItem()
        {
            string val = inputEdit.Text.Trim();
            if (val.Length == 0)
                return;
            var items = CurrentItems;
            if (editingIndex < 0)
            {
                items.Add(val);
                RefreshListBox(items.Count);
            }
            else
            {
                items[editingIndex] = val;
                RefreshListBox(editingIndex + 1);
                editingIndex = -1;
            }
            inputEdit.Text = "";
            btnAddOK.Text = "リストへ追加";
            inputEdit.Focus();
        }

        private void DeleteItem()
        {
            int idx = itemList.SelectedIndex;
            if (idx < 0)
                return;
            var items = CurrentItems;
            items.RemoveAt(idx);
            RefreshListBox(Math.Min(idx + 1, items.Count));
        }

        private void MoveItem(int direction)
        {
            int idx = itemList.SelectedIndex;
            var items = CurrentItems;
            int newIdx = idx + direction;
            if (idx < 0 || newIdx < 0 || newIdx >= items.Count)
                return;
            (items[idx], items[newIdx]) = (items[newIdx], items[idx]);
            RefreshListBox(newIdx + 1);
        }

        // 値欄に未確定の入力が残っていたら確認する
        private void ConfirmPendingInput()
        {
            var cat = categories[selectedCategoryIndex];
            if (cat.Type != "list")
                return;
            string val = inputEdit.Text.Trim();
            if (val.Length == 0)
                return;
            string action = editingIndex < 0 ? "追加" : "更新";
            var result = MessageBox.Show(
                $"値欄に入力が残っています：\n\n    {val}\n\nリストに{action}しますか？",
                "未確定の入力", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
                CommitItem();
            else
            {
                inputEdit.Text = "";
                btnAddOK.Text = "リストへ追加";
                editingIndex = -1;
            }
        }

        private bool SaveAll()
        {
            ConfirmPendingInput();

            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            for (int i = 0; i < categories.Length; i++)
            {
                var cat = categories[i];
                if (cat.Type == "list")
                {
                    string[] arr = categoryItems[i].Distinct().ToArray();
                    List[cat.ListIndex] = arr;
                    config.AppSettings.Settings[cat.AppSettingKey].Value = string.Join(",", arr);
                }
            }

            int seconds = (int)secondsCtrl.Value;
            noKeyInputInterval = seconds * 1000;
            config.AppSettings.Settings["NoKeyInputTime"].Value = (seconds / 60.0).ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);

            AppExitEndEnabled = chkAppExitEnd.Checked;
            config.AppSettings.Settings["AppExitEnd"].Value = AppExitEndEnabled ? "1" : "0";

            darkModeEnabled = chkDarkMode.Checked;
            config.AppSettings.Settings["IsDarkMode"].Value = darkModeEnabled ? "on" : "off";

            // MS-IME入力設定(設定ファイルに項目が無い場合は追加する)
            SetAppSetting(config, "KeyAssignment", chkKeyAssign.Checked ? "1" : "0");
            SetAppSetting(config, "MsImeSwitch", chkMsImeSwitch.Checked ? "1" : "0");
            SetAppSetting(config, "SpaceWidth", $"{cmbSpaceInit.SelectedIndex},{cmbSpaceTo.SelectedIndex}");
            SetAppSetting(config, "Punctuation", $"{cmbPunctInit.SelectedIndex},{cmbPunctTo.SelectedIndex}");

            config.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection("appSettings");

            // レジストリとトレイメニューへ即時反映する
            ApplyMsImeSettings();
            Debug.WriteLine("設定を保存しました");
            return true;
        }
    }
    private void MenuReload_Click(object sender, EventArgs e)
    {
        var map = new ExeConfigurationFileMap { ExeConfigFilename = "AlwaysIME.dll.config" };
        string configFilePath = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None).FilePath;
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
            Trace.WriteLine("Failed to write registory.");
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
            Trace.WriteLine("Failed to write registory.");
        }
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
