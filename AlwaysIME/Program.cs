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
    static void Main()
    {
        bool createdNew;
        Mutex mutex = new Mutex(true, "AlwaysIME", out createdNew);

        if (createdNew)
        {
            ApplicationConfiguration.Initialize();
            ResidentTest.InitializeAppConfig();
            ResidentTest rm = new ResidentTest();
            Application.Run();
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
    static string RegistrationprocessName;
    static bool ImeModeGlobal = true;
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
    static readonly string[][] List = new string[5][];
    const int PassArray = 0;
    const int ImeOffArray = 1;
    const int ImeOffTitleArray = 2;
    const int OnActivatedAppArray = 3;
    const int BackgroundArray = 4;
    private static string RunOnActivatedAppPath;
    private static string RunOnActivatedArgv;
    private static string RunBackgroundAppPath;
    private static string FWBackgroundArgv;
    private static int imeInterval = 500;
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

    static readonly int[][] val = new int[3][];
    const int ConfigPunctuation = 0;
    const int ConfigSpaceWidth = 1;
    static int SetPunctuationMode;
    static int SetSpaceMode;
    static readonly string[] keyPath = new string[2];
    static readonly string[] valueName = new string[2];
    static readonly RegistryValueKind[] valueType = new RegistryValueKind[2];
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
        this.setComponents();
        this.timer = new System.Windows.Forms.Timer();
        this.timer.Interval = imeInterval;
        this.timer.Tick += new EventHandler(Timer_Tick);
        this.timer.Start();
        this.lastInputTime = DateTime.Now;
    }

    public static void InitializeAppConfig()
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
        buff = ConfigurationManager.AppSettings["Punctuation"];
        if (!string.IsNullOrEmpty(buff))
        {
            int temp = 0x7FFCFFFF & (int)ReadRegistryValue(RegistryHive.CurrentUser, keyPath[ConfigPunctuation], valueName[ConfigPunctuation], valueType[ConfigPunctuation]);
            string[] parts = buff.Split(',');
            val[ConfigPunctuation] = new int[2];
            for (int i = 0; i < parts.Length; i++)
            {
                if (int.Parse(parts[i]) >= 0 && int.Parse(parts[i]) <= 3)
                {
                    val[ConfigPunctuation][i] = temp | int.Parse(parts[i]) << 16;
                }
                else
                {
                    MessageBox.Show("Punctuationの設定が間違っています");
                    val[ConfigPunctuation] = [0, 0];
                }
            }
            SetPunctuationMode = val[ConfigPunctuation][0];
        }
        buff = ConfigurationManager.AppSettings["SpaceWidth"];
        if (!string.IsNullOrEmpty(buff))
        {
            string[] parts = buff.Split(',');
            val[ConfigSpaceWidth] = new int[2];
            for (int i = 0; i < parts.Length; i++)
            {
                if (int.Parse(parts[i]) >= 0 && int.Parse(parts[i]) <= 2)
                {
                    val[ConfigSpaceWidth][i] = int.Parse(parts[i]);
                }
                else
                {
                    MessageBox.Show("SpaceWidthの設定が間違っています");
                    val[ConfigSpaceWidth] = [0, 0];
                }
            }
            SetSpaceMode = val[ConfigSpaceWidth][0];
        }
        buff = ConfigurationManager.AppSettings["AppExitEnd"];
        if (!string.IsNullOrEmpty(buff))
        {
            if (int.Parse(buff) == 1)
            {
                AppExitEndEnabled = true;
            }
            else/* if (int.Parse(buff) == 0)*/
            {
                AppExitEndEnabled = false;
            }
        }
        else
        {
            AppExitEndEnabled = true;
        }
    }
    private void Close_Click(object sender, EventArgs e)
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
        icon.Visible = false;
        icon.Dispose();
        Application.Exit();
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
        menuItemModeOn.Click += new EventHandler((sender, e) => ChangeAlwaysIMEModeAndSave("on"));
        ToolStripMenuItem menuItemModeOff = new ToolStripMenuItem("グローバル");
        menuItemModeOff.Click += new EventHandler((sender, e) => ChangeAlwaysIMEModeAndSave("off"));

        ToolStripSeparator separator2 = new ToolStripSeparator();

        ToolStripMenuItem updateTimeMenuItem = new ToolStripMenuItem("更新間隔");
        ToolStripMenuItem menuItem250 = new ToolStripMenuItem("250 ms");
        menuItem250.Click += new EventHandler((sender, e) => ChangeIntervalAndSave(250));
        ToolStripMenuItem menuItem500 = new ToolStripMenuItem("500 ms");
        menuItem500.Click += new EventHandler((sender, e) => ChangeIntervalAndSave(500));
        ToolStripMenuItem menuItem1000 = new ToolStripMenuItem("1000 ms");
        menuItem1000.Click += new EventHandler((sender, e) => ChangeIntervalAndSave(1000));

        ToolStripSeparator separator3 = new ToolStripSeparator();

        ToolStripMenuItem MenuItemRegistrationDialog = new ToolStripMenuItem();
        MenuItemRegistrationDialog.Text = "IMEオフに登録";
        MenuItemRegistrationDialog.Click += MenuItemRegistrationDialog_Click;

        ToolStripSeparator separator4 = new ToolStripSeparator();

        ToolStripMenuItem menuPunctuation = new ToolStripMenuItem();
        menuPunctuation.Text = "句読点切替(&P)";
        menuPunctuation.Click += new EventHandler(Punctuation_Click);

        ToolStripSeparator separator5 = new ToolStripSeparator();

        ToolStripMenuItem menuSpace = new ToolStripMenuItem();
        menuSpace.Text = "スペース切替(&S)";
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
            menuItem250.BackColor = Color.FromArgb(32, 32, 32);
            menuItem250.ForeColor = Color.White;
            menuItem500.BackColor = Color.FromArgb(32, 32, 32);
            menuItem500.ForeColor = Color.White;
            menuItem1000.BackColor = Color.FromArgb(32, 32, 32);
            menuItem1000.ForeColor = Color.White;
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
            MenuItemRegistrationDialog.BackColor = Color.FromArgb(32, 32, 32);
            MenuItemRegistrationDialog.ForeColor = Color.White;
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
        updateTimeMenuItem.DropDownItems.Add(menuItem250);
        updateTimeMenuItem.DropDownItems.Add(menuItem500);
        updateTimeMenuItem.DropDownItems.Add(menuItem1000);
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
        menu.Items.Add(MenuItemRegistrationDialog);
        if (val[ConfigPunctuation][0] != val[ConfigPunctuation][1])
        {
            if (!darkModeEnabled)
                menu.Items.Add(separator4);
            menu.Items.Add(menuPunctuation);
        }
        if (val[ConfigSpaceWidth][0] != val[ConfigSpaceWidth][1])
        {
            if (!darkModeEnabled)
                menu.Items.Add(separator5);
            menu.Items.Add(menuSpace);
        }
        if (AppExitEndEnabled)
        {
            if (!darkModeEnabled)
                menu.Items.Add(separator6);
            menu.Items.Add(menuItem);
        }
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
            float Zoom;
            using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
                Zoom = graphics.DpiX / 96;
            this.Font = new Font("Meiryo", (int)(9 * Math.Pow(Zoom, 1.0 / 3.0)), FontStyle.Regular);
            this.Text = "IMEオフに登録";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new Size((int)(440 * Zoom), (int)(200 * Zoom));

            titleLabel = new Label();
            titleLabel.Text = "タイトル:";
            titleLabel.Size = new Size((int)(80 * Zoom), (int)(20 * Zoom));
            titleLabel.Location = new Point((int)(20 * Zoom), (int)(20 * Zoom));
            titleTextBox = new TextBox();
            titleTextBox.Location = new Point((int)(100 * Zoom), (int)(20 * Zoom));
            titleTextBox.Size = new Size((int)(300 * Zoom), (int)(20 * Zoom));
            titleTextBox.Text = RegistrationWindowTitle;
            appLabel = new Label();
            appLabel.Text = "アプリ名:";
            appLabel.Size = new Size((int)(80 * Zoom), (int)(20 * Zoom));
            appLabel.Location = new Point((int)(20 * Zoom), (int)(50 * Zoom));
            appTextBox = new TextBox();
            appTextBox.Location = new Point((int)(100 * Zoom), (int)(50 * Zoom));
            appTextBox.Size = new Size((int)(300 * Zoom), (int)(20 * Zoom));
            appTextBox.Text = RegistrationprocessName;
            titleRadioButton = new RadioButton();
            titleRadioButton.Text = "タイトル";
            titleRadioButton.Size = new Size((int)(90 * Zoom), (int)(30 * Zoom));
            titleRadioButton.Location = new Point((int)(100 * Zoom), (int)(80 * Zoom));
            appRadioButton = new RadioButton();
            appRadioButton.Text = "アプリ名";
            appRadioButton.Size = new Size((int)(90 * Zoom), (int)(30 * Zoom));
            appRadioButton.Location = new Point((int)(220 * Zoom), (int)(80 * Zoom));
            appRadioButton.Checked = true;
            okButton = new Button();
            okButton.Text = "登録(&R)";
            okButton.Size = new Size((int)(110 * Zoom), (int)(32 * Zoom));
            okButton.DialogResult = DialogResult.OK;
            okButton.Location = new Point((int)(100 * Zoom), (int)(110 * Zoom));
            okButton.Click += OkButton_Click;
            cancelButton = new Button();
            cancelButton.Text = "キャンセル(&C)";
            cancelButton.Size = new Size((int)(110 * Zoom), (int)(32 * Zoom));
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.Location = new System.Drawing.Point((int)(220 * Zoom), (int)(110 * Zoom));

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
                List[ImeOffTitleArray] = list.ToArray();
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
                List[ImeOffArray] = list.ToArray();
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
    private void Punctuation_Click(object sender, EventArgs e)
    {
        int newValue;
        if (SetPunctuationMode == val[ConfigPunctuation][0])
        {
            newValue = val[ConfigPunctuation][1];
        }
        else
        {
            newValue = val[ConfigPunctuation][0];
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
        }
        else
        {
            newValue = val[ConfigSpaceWidth][0];
        }
        if (WriteRegistryValue(RegistryHive.CurrentUser, keyPath[ConfigSpaceWidth], valueName[ConfigSpaceWidth], newValue, valueType[ConfigSpaceWidth]))
        {
            switch (newValue)
            {
                case IME_AUTO_WIDTH_SPACE:
                    Debug.WriteLine($"スペースを現在の入力モードにしました");
                    break;
                case IME_FULL_WIDTH_SPACE:
                    Debug.WriteLine($"スペースを常に全角にしました");
                    break;
                case IME_HALF_WIDTH_SPACE:
                    Debug.WriteLine($"スペースを常に半角にしました");
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
    private void ChangeAlwaysIMEModeAndSave(String mode)
    {
        // App.config の更新
        Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        config.AppSettings.Settings["AlwaysIMEMode"].Value = mode;
        config.Save(ConfigurationSaveMode.Modified);
        ConfigurationManager.RefreshSection("appSettings");

        if (mode.CompareTo("off") == 0)
        {
            ImeModeGlobal = true;
        }
        else
        {
            ImeModeGlobal = false;
        }
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
