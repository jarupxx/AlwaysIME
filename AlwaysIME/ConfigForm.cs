using System.Configuration;
using System.Diagnostics;
using System.Globalization;

// ============================================================
// 設定ウィンドウ
// カテゴリー一覧（左） + 選択カテゴリーの編集パネル（右）の
// 2ペイン構成。リスト系設定はその場で追加・編集・削除・並べ替えができる。
// ============================================================
[System.ComponentModel.DesignerCategory("Code")]
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
            ListIndex = ResidentTest.PassArray,
        },
        new ConfigCategory
        {
            AppSettingKey = "ImeOffList",
            Label = "IME-OFFアプリ",
            Desc = "アプリの切り替えやウィンドウタイトルの変化時に、IMEをOFFにするアプリです。\r\n実行ファイル名（拡張子なし）を1行1件で入力してください。\r\n例: cmd, powershell",
            Type = "list",
            ListIndex = ResidentTest.ImeOffArray,
        },
        new ConfigCategory
        {
            AppSettingKey = "ImeOffTitle",
            Label = "IME-OFFタイトルパターン",
            Desc = "ウィンドウタイトルにマッチしたら、IMEをOFFにする正規表現です。\r\n1行1パターン、全アプリ共通で適用されます。\r\n例: \\.cs$",
            Type = "list",
            ListIndex = ResidentTest.ImeOffTitleArray,
        },
        new ConfigCategory
        {
            AppSettingKey = "UpdateTag",
            Label = "タイトル除外タグ",
            Desc = "ウィンドウタイトルの変化を判定する際に、無視する部分を表す正規表現です。\r\nマッチした部分を除去してから変化を判定します。\r\n例: \\(更新\\)",
            Type = "list",
            ListIndex = ResidentTest.RemoveUpdateTagArray,
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
        new ConfigCategory
        {
            AppSettingKey = null,
            Label = "バックアップ",
            Desc = "設定ファイルのバックアップと復元を行います。\r\n保存先は設定ファイルと同じ場所の backup フォルダーです。\r\nバックアップの対象は保存済みの設定です。",
            Type = "backup",
        },
        new ConfigCategory
        {
            AppSettingKey = null,
            Label = "拡張設定",
            Desc = "AlwaysIME.dll.config の全設定をここで変更できます。\r\n設定内容の詳細が分からないまま変更しないでください。\r\n「リセット」で初期値に戻せます（保存して閉じる／適用で反映）。",
            Type = "extended",
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

    // バックアップパネル
    private static readonly double[] BackupHoursChoices = [1, 6, 12, 24, 48, 72];
    private static readonly int[] BackupKeepChoices = [3, 5, 10, 20, 50];
    private static readonly string[] RestoreScopeLabels = ["すべての設定", "アプリ一覧・連携設定のみ", "それ以外の設定のみ"];
    // リンクの色（通常モード / ダークモードそれぞれの背景で読みやすい緑）
    private static readonly Color BackupLinkColor = Color.FromArgb(26, 127, 55);
    private static readonly Color BackupLinkColorDark = Color.FromArgb(63, 185, 80);
    private ComboBox cmbBackupHours, cmbBackupKeep, cmbRestoreDate, cmbRestoreScope;
    private readonly List<double> backupHoursValues = new List<double>();
    private readonly List<int> backupKeepValues = new List<int>();
    private List<BackupService.BackupEntry> backupChoices = new List<BackupService.BackupEntry>();
    private LinkLabel linkBackupFolder, linkBackupNow;
    private Label lblBackupLatest;
    private Button btnRestore;
    private Font backupItalicFont;
    private Control[] backupControls;

    // 拡張設定パネル（キー / 値 / リセット）
    private const string ExtColKey = "colKey";
    private const string ExtColValue = "colValue";
    private const string ExtColReset = "colReset";
    private TextBox extFilter;
    private DataGridView extGrid;
    private Font extBoldFont;

    // 拡張設定の各キーの説明（セルにマウスを合わせたときのポップアップ表示）
    private static readonly Dictionary<string, string> ConfigDescriptions = new Dictionary<string, string>
    {
        { "AlwaysIMEMode", "IMEの制御モードです。\n1: IMEオン（常にオンにする）\n2: IMEオフ\n3: グローバル（直前のIME状態を維持する）\n※ on / off も指定できます（on=1、off=3）。" },
        { "IsDarkMode", "トレイメニューや設定画面をダークモードにします。\non: ダークモード / それ以外: 通常\n※ メニューへの反映は再起動後です。" },
        { "PassList", "IMEを一切操作しないアプリです。\n実行ファイル名（拡張子なし）をカンマ区切りで指定します。\n例: AlwaysIME,notepad" },
        { "ImeOffList", "アプリの切り替えやウィンドウタイトルの変化時に、IMEをオフにするアプリです。\n実行ファイル名（拡張子なし）をカンマ区切りで指定します。\n例: cmd,powershell" },
        { "ImeOffTitle", "ウィンドウタイトルにマッチしたらIMEをオフにする正規表現です。\nカンマ区切りで複数指定でき、全アプリ共通で適用されます。\n例: \\.cs$" },
        { "UpdateTag", "ウィンドウタイトルの変化を判定するときに無視する部分の正規表現です。\nマッチした部分を除去してから変化を判定します。カンマ区切りで複数指定できます。\n例: \\(更新\\)" },
        { "intervalTime", "IMEの状態を監視する間隔です（ミリ秒）。\n小さいほど反応が早くなりますが、負荷が増えます。\n例: 200" },
        { "SuspendFewTime", "トレイメニュー「少し無効」で、機能を停止する時間です（分）。" },
        { "SuspendTime", "トレイメニュー「しばらく無効」で、機能を停止する時間です（分）。" },
        { "NoKeyInputTime", "この時間、キーボードやマウスの入力が無かったら、IMEの制御状態を復元します。\n単位は分で、小数も指定できます（0.1 = 6秒）。" },
        { "DelayBackgroundTime", "BackgroundAppList のアプリが非アクティブになってから、連携アプリ（BackgroundAppPath）を起動するまでの時間です（分）。\n空欄の場合は無効です。" },
        { "OnActivatedAppList", "アクティブになったときに連携アプリ（OnActivatedAppPath）を起動するアプリです。\n実行ファイル名（拡張子なし）をカンマ区切りで指定します。" },
        { "OnActivatedAppPath", "OnActivatedAppList のアプリがアクティブになったときに起動する連携アプリのパスです。\n存在しないパスを指定すると警告が表示されます。" },
        { "OnActivatedArgv", "OnActivatedAppPath の連携アプリに渡すコマンドライン引数です。" },
        { "BackgroundAppList", "非アクティブになったときに連携アプリ（BackgroundAppPath）の起動を予約するアプリです。\n実行ファイル名（拡張子なし）をカンマ区切りで指定します。" },
        { "BackgroundAppPath", "BackgroundAppList のアプリが非アクティブになり、DelayBackgroundTime が経過したときに起動する連携アプリのパスです。\n存在しないパスを指定すると警告が表示されます。" },
        { "BackgroundArgv", "BackgroundAppPath の連携アプリに渡すコマンドライン引数です。" },
        { "Punctuation", "MS-IMEの句読点を「初期値,切替先」の形式で指定します。\n0: ，．  1: 、。  2: 、．  3: ，。\n例: 1,3\n※ MsImeSwitch が 1 のときに有効です。" },
        { "SpaceWidth", "MS-IMEのスペース入力を「初期値,切替先」の形式で指定します。\n0: 現在の入力モード  1: 常に全角  2: 常に半角\n例: 1,2\n※ MsImeSwitch が 1 のときに有効です。" },
        { "MsImeSwitch", "スペース / 句読点の切替を有効にするかどうかです。\n1: 有効（トレイメニューから切替できます）\n0: 無効（レジストリへの書き込みを行いません）" },
        { "KeyAssignment", "変換キー→IME-ON / 無変換キー→IME-OFF の割り当てです。\n1: 割り当てる  0: 割り当てない\n空欄: 管理しない（終了時には変更前の設定に戻します）" },
        { "AppExitEnd", "トレイメニューの「常駐の終了」の位置です。\n1: メニューの末尾に配置  0: 上部に配置\n※ 反映は再起動後です。" },
        { "BackupStableHours", "設定を変更してから、アプリがこの時間（単位: 時間）動作し続けたら、安定動作を確認できたとみなして「安定版バックアップ」を作成します。\n設定を試行錯誤している間は作成されません。\n0 を指定すると、確認のたびに作成します。" },
        { "BackupKeep", "保持する安定版バックアップの世代数です（1〜50）。\n超えた分は古いものから削除します。\n保存先: 設定ファイルと同じ場所の backup フォルダー" },
    };
    private Control[] extControls;
    private bool extLoading = false;
    // 読み込み時点（保存済み）の値 / 未保存の変更
    private readonly Dictionary<string, string> extBaseline = new Dictionary<string, string>();
    private readonly Dictionary<string, string> extPending = new Dictionary<string, string>();

    private Button btnSave, btnApply, btnCancel;

    public ConfigForm()
    {
        InitializeComponents();
        LoadCurrentValues();
        LoadExtendedSettings();
        RenderCategory(0);
    }

    private void InitializeComponents()
    {
        using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
            uiScale = g.DpiX / 96;

        int W = (int)(720 * uiScale);
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

        if (ResidentTest.darkModeEnabled)
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

        // ---- 拡張設定パネル（キー / 値 / リセット） ----
        extFilter = new TextBox
        {
            Location = new Point(RX, panelTop),
            Size = new Size(RW, (int)(24 * uiScale)),
            PlaceholderText = "フィルタ設定",
        };
        extFilter.TextChanged += (s, e) => ApplyExtFilter();

        extGrid = new DataGridView
        {
            Location = new Point(RX, panelTop + (int)(28 * uiScale)),
            Size = new Size(RW, panelHeight - (int)(28 * uiScale)),
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            RowHeadersVisible = false,
            MultiSelect = false,
            SelectionMode = DataGridViewSelectionMode.CellSelect,
            EditMode = DataGridViewEditMode.EditOnEnter,
            BackgroundColor = SystemColors.Window,
            BorderStyle = BorderStyle.FixedSingle,
        };
        extGrid.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = ExtColKey,
            HeaderText = "キー",
            ReadOnly = true,
            Width = (int)(150 * uiScale),
            SortMode = DataGridViewColumnSortMode.Automatic,
        });
        extGrid.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = ExtColValue,
            HeaderText = "値",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            SortMode = DataGridViewColumnSortMode.NotSortable,
        });
        extGrid.Columns.Add(new DataGridViewButtonColumn
        {
            Name = ExtColReset,
            HeaderText = "",
            Text = "リセット",
            UseColumnTextForButtonValue = true,
            FlatStyle = FlatStyle.Flat,
            Width = (int)(76 * uiScale),
            SortMode = DataGridViewColumnSortMode.NotSortable,
        });
        extBoldFont = new Font(extGrid.Font, FontStyle.Bold);
        extGrid.CellContentClick += ExtGrid_CellContentClick;
        extGrid.CellEndEdit += ExtGrid_CellEndEdit;
        extGrid.DataError += (s, e) => e.ThrowException = false;

        if (ResidentTest.darkModeEnabled)
        {
            extGrid.BackgroundColor = Color.FromArgb(32, 32, 32);
            extGrid.GridColor = Color.FromArgb(70, 70, 70);
            extGrid.EnableHeadersVisualStyles = false;
            extGrid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(55, 55, 55);
            extGrid.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            extGrid.DefaultCellStyle.BackColor = Color.FromArgb(45, 45, 45);
            extGrid.DefaultCellStyle.ForeColor = Color.White;
            extGrid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(80, 80, 120);
            extGrid.DefaultCellStyle.SelectionForeColor = Color.White;
        }

        extControls = new Control[] { extFilter, extGrid };
        this.Controls.AddRange(extControls);

        // ---- バックアップパネル ----
        int bkComboW = (int)(RW * 0.42);
        int bkComboX = RX + RW - bkComboW;
        int bkRow1 = panelTop + (int)(28 * uiScale);
        int bkRow2 = panelTop + (int)(56 * uiScale);
        int bkLblDy = (int)(4 * uiScale);
        int bkLblW = RW - bkComboW - SP;

        Label hdrAuto = new Label
        {
            Text = "自動バックアップ",
            Location = new Point(RX, panelTop),
            Size = new Size(RW, (int)(20 * uiScale)),
            Font = extBoldFont,
        };
        Label lblBackupHours = MakeMsImeLabel("安定版の作成タイミング", RX, bkRow1 + bkLblDy, bkLblW, (int)(20 * uiScale));
        cmbBackupHours = MakeMsImeCombo(Array.Empty<string>(), bkComboX, bkRow1, bkComboW);
        Label lblBackupKeep = MakeMsImeLabel("バックアップの最大数", RX, bkRow2 + bkLblDy, bkLblW, (int)(20 * uiScale));
        cmbBackupKeep = MakeMsImeCombo(Array.Empty<string>(), bkComboX, bkRow2, bkComboW);

        Color backupLinkColor = ResidentTest.darkModeEnabled ? BackupLinkColorDark : BackupLinkColor;
        linkBackupFolder = new LinkLabel
        {
            Text = "バックアップフォルダーを表示📂",
            Location = new Point(RX, panelTop + (int)(92 * uiScale)),
            AutoSize = true,
            LinkColor = backupLinkColor,
            ActiveLinkColor = backupLinkColor,
            VisitedLinkColor = backupLinkColor,
            LinkBehavior = LinkBehavior.HoverUnderline,
        };
        linkBackupFolder.LinkClicked += (s, e) => OpenBackupFolder();

        linkBackupNow = new LinkLabel
        {
            Text = "今すぐバックアップ💾",
            Location = new Point(RX, panelTop + (int)(116 * uiScale)),
            AutoSize = true,
            LinkColor = backupLinkColor,
            ActiveLinkColor = backupLinkColor,
            VisitedLinkColor = backupLinkColor,
            LinkBehavior = LinkBehavior.HoverUnderline,
        };
        linkBackupNow.LinkClicked += (s, e) => BackupNowClicked();

        backupItalicFont = new Font(this.Font, FontStyle.Italic);
        lblBackupLatest = new Label
        {
            Location = new Point(RX, panelTop + (int)(138 * uiScale)),
            Size = new Size(RW, (int)(36 * uiScale)),
            Font = backupItalicFont,
        };

        Label backupSeparator = new Label
        {
            BorderStyle = BorderStyle.Fixed3D,
            Location = new Point(RX, panelTop + (int)(180 * uiScale)),
            Size = new Size(RW, 2),
        };

        Label hdrRestore = new Label
        {
            Text = "バックアップの復元",
            Location = new Point(RX, panelTop + (int)(190 * uiScale)),
            Size = new Size(RW, (int)(20 * uiScale)),
            Font = extBoldFont,
        };
        int restoreW = (int)(RW * 0.7);
        cmbRestoreDate = MakeMsImeCombo(Array.Empty<string>(), RX, panelTop + (int)(216 * uiScale), restoreW);
        cmbRestoreScope = MakeMsImeCombo(Array.Empty<string>(), RX, panelTop + (int)(244 * uiScale), restoreW);
        cmbRestoreScope.Items.Add("復元オプションを選択");
        cmbRestoreScope.Items.AddRange(RestoreScopeLabels);
        cmbRestoreScope.SelectedIndex = 0;
        btnRestore = new Button
        {
            Text = "復元する",
            Enabled = false,
            Location = new Point(RX, panelTop + (int)(274 * uiScale)),
            Size = new Size(restoreW, (int)(26 * uiScale)),
        };
        cmbRestoreDate.SelectedIndexChanged += (s, e) => UpdateRestoreButton();
        cmbRestoreScope.SelectedIndexChanged += (s, e) => UpdateRestoreButton();
        btnRestore.Click += (s, e) => RestoreSelectedBackup();

        backupControls = new Control[]
        {
            hdrAuto, lblBackupHours, cmbBackupHours, lblBackupKeep, cmbBackupKeep,
            linkBackupFolder, linkBackupNow, lblBackupLatest, backupSeparator,
            hdrRestore, cmbRestoreDate, cmbRestoreScope, btnRestore,
        };
        this.Controls.AddRange(backupControls);

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

        if (ResidentTest.darkModeEnabled)
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

        // バックアップ画面の補足の色（ダークモードの一括着色より後に設定する）
        lblBackupLatest.ForeColor = Color.Gray;
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
                string[] source = ResidentTest.List[cat.ListIndex];
                categoryItems[Array.IndexOf(categories, cat)] = source != null
                    ? new List<string>(source)
                    : new List<string>();
            }
        }
        LoadBackupSettings();
        secondsCtrl.Value = Math.Max(secondsCtrl.Minimum, Math.Min(secondsCtrl.Maximum, ResidentTest.noKeyInputInterval / 1000));
        chkAppExitEnd.Checked = ResidentTest.AppExitEndEnabled;
        chkDarkMode.Checked = ResidentTest.darkModeEnabled;

        // ---- MS-IME入力設定 ----
        // キー割り当て: 設定済みならその値、未設定ならレジストリの現在値
        string keyAssignBuff = ConfigurationManager.AppSettings["KeyAssignment"];
        chkKeyAssign.Checked = !string.IsNullOrEmpty(keyAssignBuff)
            ? keyAssignBuff == "1"
            : ResidentTest.ReadMsImeDword("IsKeyAssignmentEnabled") == 1;

        // 切替の有効/無効: 明示設定があればその値、未設定なら従来の設定値の有無で判定
        string switchBuff = ConfigurationManager.AppSettings["MsImeSwitch"];
        string punctBuff = ConfigurationManager.AppSettings["Punctuation"];
        string spaceBuff = ConfigurationManager.AppSettings["SpaceWidth"];
        chkMsImeSwitch.Checked = string.IsNullOrEmpty(switchBuff)
            ? (ResidentTest.IsMsImePairConfigured(punctBuff) || ResidentTest.IsMsImePairConfigured(spaceBuff))
            : switchBuff == "1";

        // 初期値/切替先: 設定済みならその値、未設定ならレジストリの現在値を初期値にする
        int regSpace = ResidentTest.ReadMsImeDword("InputSpace") ?? ResidentTest.IME_AUTO_WIDTH_SPACE;
        if (regSpace < ResidentTest.IME_AUTO_WIDTH_SPACE || regSpace > ResidentTest.IME_HALF_WIDTH_SPACE)
            regSpace = ResidentTest.IME_AUTO_WIDTH_SPACE;
        int? regOption1 = ResidentTest.ReadMsImeDword("option1");
        int regPunct = regOption1.HasValue ? (regOption1.Value >> 16) & 0x3 : ResidentTest.IME_TOUTEN_KUTENN;

        int spaceInit = regSpace;
        int spaceTo = regSpace == ResidentTest.IME_HALF_WIDTH_SPACE ? ResidentTest.IME_FULL_WIDTH_SPACE : ResidentTest.IME_HALF_WIDTH_SPACE;
        if (ResidentTest.TryParseConfigPair(spaceBuff, ResidentTest.IME_HALF_WIDTH_SPACE, out int cfgSpaceInit, out int cfgSpaceTo)
            && !(cfgSpaceInit == 0 && cfgSpaceTo == 0))
        {
            spaceInit = cfgSpaceInit;
            spaceTo = cfgSpaceTo;
        }

        int punctInit = regPunct;
        int punctTo = ResidentTest.IME_COMMA_KUTENN;
        if (ResidentTest.TryParseConfigPair(punctBuff, ResidentTest.IME_COMMA_KUTENN, out int cfgPunctInit, out int cfgPunctTo)
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

    // ============================================================
    // バックアップ（自動バックアップの設定 / 今すぐバックアップ / 復元）
    // ============================================================

    // 設定ファイルの値を、コンボボックスへ読み込む（一覧にない値は選択肢として追加する）
    private void LoadBackupSettings()
    {
        double hours = 24;
        if (double.TryParse(ConfigurationManager.AppSettings["BackupStableHours"], NumberStyles.Float,
                CultureInfo.InvariantCulture, out double h) && h >= 0 && !double.IsInfinity(h))
            hours = h;
        var hoursList = new List<double>(BackupHoursChoices);
        if (!hoursList.Contains(hours))
        {
            hoursList.Add(hours);
            hoursList.Sort();
        }
        backupHoursValues.Clear();
        cmbBackupHours.Items.Clear();
        foreach (double v in hoursList)
        {
            backupHoursValues.Add(v);
            cmbBackupHours.Items.Add(v == 0 ? "確認のたびに作成" : $"変更後 {v:0.##}時間動作したら");
        }
        cmbBackupHours.SelectedIndex = backupHoursValues.IndexOf(hours);

        int keep = 5;
        if (int.TryParse(ConfigurationManager.AppSettings["BackupKeep"], NumberStyles.Integer,
                CultureInfo.InvariantCulture, out int k))
            keep = Math.Clamp(k, 1, 50);
        var keepList = new List<int>(BackupKeepChoices);
        if (!keepList.Contains(keep))
        {
            keepList.Add(keep);
            keepList.Sort();
        }
        backupKeepValues.Clear();
        cmbBackupKeep.Items.Clear();
        foreach (int v in keepList)
        {
            backupKeepValues.Add(v);
            cmbBackupKeep.Items.Add(v.ToString());
        }
        cmbBackupKeep.SelectedIndex = backupKeepValues.IndexOf(keep);
    }

    private static string DescribeBackup(BackupService.BackupEntry entry)
    {
        string date = entry.SavedUtc.ToLocalTime().ToString("yyyy/MM/dd HH:mm");
        return entry.Kind switch
        {
            BackupService.BackupKind.Latest => $"最新の設定　{date}",
            BackupService.BackupKind.Stable => $"安定版 #{entry.Seq}　{date}",
            _ => $"復元前の設定 #{entry.Seq}　{date}",
        };
    }

    // 最新のバックアップの表示と、復元できる一覧を更新する
    private void RefreshBackupPanel()
    {
        backupChoices = BackupService.GetBackups();

        cmbRestoreDate.Items.Clear();
        cmbRestoreDate.Items.Add("日付を選択");
        foreach (var entry in backupChoices)
            cmbRestoreDate.Items.Add(DescribeBackup(entry));
        cmbRestoreDate.SelectedIndex = 0;
        cmbRestoreScope.SelectedIndex = 0;

        var latest = backupChoices.FirstOrDefault(b => b.Kind == BackupService.BackupKind.Latest);
        var stable = backupChoices.FirstOrDefault(b => b.Kind == BackupService.BackupKind.Stable);
        lblBackupLatest.Text =
            "最新のバックアップ: " + (latest != null ? latest.SavedUtc.ToLocalTime().ToString("yyyy年M月d日 HH:mm") : "なし") + "\r\n" +
            "最新の安定版: " + (stable != null
                ? $"#{stable.Seq}（{stable.SavedUtc.ToLocalTime():yyyy年M月d日 HH:mm}）"
                : "なし");
        UpdateRestoreButton();
    }

    private void UpdateRestoreButton()
    {
        btnRestore.Enabled = cmbRestoreDate.SelectedIndex > 0 && cmbRestoreScope.SelectedIndex > 0;
    }

    private void OpenBackupFolder()
    {
        string dir = BackupService.BackupDirectory;
        if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir))
        {
            MessageBox.Show(this, "バックアップフォルダーがありません。", "AlwaysIME", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        try
        {
            Process.Start(new ProcessStartInfo { FileName = dir, UseShellExecute = true });
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "フォルダーを開けませんでした。\n\n" + ex.Message, "AlwaysIME", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void BackupNowClicked()
    {
        BackupService.BackupNow(out string message);
        RefreshBackupPanel();
        MessageBox.Show(this, message, "バックアップ", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private void RestoreSelectedBackup()
    {
        int dateIdx = cmbRestoreDate.SelectedIndex - 1;
        int scopeIdx = cmbRestoreScope.SelectedIndex - 1;
        if (dateIdx < 0 || dateIdx >= backupChoices.Count || scopeIdx < 0)
            return;
        BackupService.BackupEntry entry = backupChoices[dateIdx];
        var scope = (BackupService.RestoreScope)scopeIdx;

        string confirm = $"「{DescribeBackup(entry)}」から復元します。\n復元する範囲: {RestoreScopeLabels[scopeIdx]}\n\n" +
            "現在の設定は「復元前の設定」として残します。\nこの画面で保存していない変更は失われます。\n\n復元しますか？";
        if (MessageBox.Show(this, confirm, "バックアップの復元", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;

        if (!BackupService.RestoreFrom(entry, scope, out string error))
        {
            MessageBox.Show(this, "復元できませんでした。\n\n" + error, "バックアップの復元", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // 復元した内容をアプリと各画面へ反映する
        ConfigurationManager.RefreshSection("appSettings");
        ResidentTest.ApplyMsImeSettings();
        LoadCurrentValues();
        LoadExtendedSettings();
        RenderCategory(selectedCategoryIndex);
        MessageBox.Show(this, "復元しました。", "バックアップの復元", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
            extBoldFont?.Dispose();
            backupItalicFont?.Dispose();
        base.Dispose(disposing);
    }

    // ============================================================
    // 拡張設定（AlwaysIME.dll.config の全キーを キー / 値 / リセット で編集）
    // ============================================================

    // AlwaysIME.dll.config の appSettings を読み込んで一覧を作り直す。
    // ファイルに無い既知のキーは初期値を表示する（編集しない限り書き込まない）。
    private void LoadExtendedSettings()
    {
        extLoading = true;
        try
        {
            extGrid.EndEdit();
            extGrid.CurrentCell = null;
            extGrid.Rows.Clear();
            extBaseline.Clear();
            extPending.Clear();

            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            foreach (string key in config.AppSettings.Settings.AllKeys)
                extBaseline[key] = config.AppSettings.Settings[key].Value ?? "";
            foreach (var kv in ResidentTest.DefaultConfigValues)
            {
                if (!extBaseline.ContainsKey(kv.Key))
                    extBaseline[kv.Key] = kv.Value;
            }

            var keys = new List<string>(extBaseline.Keys);
            keys.Sort(StringComparer.OrdinalIgnoreCase);
            foreach (string key in keys)
            {
                int row = extGrid.Rows.Add(key, extBaseline[key]);
                UpdateExtRowStyle(extGrid.Rows[row]);
                SetExtToolTips(extGrid.Rows[row], key);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("設定ファイルの読み込みに失敗しました: " + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            extLoading = false;
        }
        ApplyExtFilter();
    }

    // キー・値・リセットの各セルに、説明と初期値をポップアップで表示する
    private static void SetExtToolTips(DataGridViewRow row, string key)
    {
        bool hasDefault = ResidentTest.DefaultConfigValues.TryGetValue(key, out string defaultValue);
        string description = ConfigDescriptions.TryGetValue(key, out string d)
            ? d
            : "この項目の説明はありません。";
        string defaultText = hasDefault
            ? (defaultValue.Length == 0 ? "初期値: （空欄）" : "初期値: " + defaultValue)
            : "初期値: （なし）";

        string text = key + "\n\n" + description + "\n\n" + defaultText;
        row.Cells[ExtColKey].ToolTipText = text;
        row.Cells[ExtColValue].ToolTipText = text;
        row.Cells[ExtColReset].ToolTipText = hasDefault
            ? (defaultValue.Length == 0
                ? "「" + key + "」を初期値（空欄）に戻します。"
                : "「" + key + "」を初期値「" + defaultValue + "」に戻します。")
            : "この項目には初期値がないため、リセットできません。";
    }

    // キーまたは値にフィルタ文字列を含む行だけを表示する
    private void ApplyExtFilter()
    {
        if (extGrid == null || extLoading)
            return;
        string filter = extFilter.Text.Trim();
        extGrid.EndEdit();
        extGrid.CurrentCell = null; // 現在行は非表示にできないため
        foreach (DataGridViewRow row in extGrid.Rows)
        {
            string key = Convert.ToString(row.Cells[ExtColKey].Value) ?? "";
            string value = Convert.ToString(row.Cells[ExtColValue].Value) ?? "";
            row.Visible = filter.Length == 0
                || key.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0
                || value.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }

    // 値セルの編集が終わったら変更を記録する
    private void ExtGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
    {
        if (extLoading || e.RowIndex < 0 || extGrid.Columns[e.ColumnIndex].Name != ExtColValue)
            return;
        DataGridViewRow row = extGrid.Rows[e.RowIndex];
        SetExtValue(row, Convert.ToString(row.Cells[ExtColValue].Value) ?? "");
    }

    // 「リセット」ボタン: そのキーの値を初期値に戻す
    private void ExtGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || extGrid.Columns[e.ColumnIndex].Name != ExtColReset)
            return;
        DataGridViewRow row = extGrid.Rows[e.RowIndex];
        string key = Convert.ToString(row.Cells[ExtColKey].Value) ?? "";
        if (!ResidentTest.DefaultConfigValues.TryGetValue(key, out string defaultValue))
            return;
        extGrid.EndEdit();
        SetExtValue(row, defaultValue);
    }

    private void SetExtValue(DataGridViewRow row, string value)
    {
        string key = Convert.ToString(row.Cells[ExtColKey].Value) ?? "";
        extLoading = true;
        row.Cells[ExtColValue].Value = value;
        extLoading = false;

        if (extBaseline.TryGetValue(key, out string baseline) && baseline == value)
            extPending.Remove(key);
        else
            extPending[key] = value;
        UpdateExtRowStyle(row);
    }

    // 変更済みの値は太字+背景色、初期値のままならリセットボタンを灰色にする
    private void UpdateExtRowStyle(DataGridViewRow row)
    {
        string key = Convert.ToString(row.Cells[ExtColKey].Value) ?? "";
        string current = Convert.ToString(row.Cells[ExtColValue].Value) ?? "";
        bool changed = extPending.ContainsKey(key);
        bool atDefault = !ResidentTest.DefaultConfigValues.TryGetValue(key, out string defaultValue) || current == defaultValue;

        DataGridViewCell valueCell = row.Cells[ExtColValue];
        valueCell.Style.Font = changed ? extBoldFont : null;
        valueCell.Style.BackColor = changed
            ? (ResidentTest.darkModeEnabled ? Color.FromArgb(75, 65, 25) : Color.FromArgb(255, 250, 205))
            : Color.Empty;

        DataGridViewCell resetCell = row.Cells[ExtColReset];
        resetCell.Style.ForeColor = atDefault
            ? Color.Gray
            : (ResidentTest.darkModeEnabled ? Color.White : SystemColors.ControlText);
        resetCell.Style.BackColor = ResidentTest.darkModeEnabled ? Color.FromArgb(55, 55, 55) : SystemColors.Control;
    }

    private void CatList_SelectedIndexChanged(object sender, EventArgs e)
    {
        int idx = catList.SelectedIndex;
        if (idx < 0)
            return;
        extGrid.EndEdit();
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

        foreach (Control c in extControls)
            c.Visible = cat.Type == "extended";

        bool isBackup = cat.Type == "backup";
        foreach (Control c in backupControls)
            c.Visible = isBackup;
        if (isBackup)
            RefreshBackupPanel();

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
                ResidentTest.List[cat.ListIndex] = arr;
                config.AppSettings.Settings[cat.AppSettingKey].Value = string.Join(",", arr);
            }
        }

        int seconds = (int)secondsCtrl.Value;
        ResidentTest.noKeyInputInterval = seconds * 1000;
        config.AppSettings.Settings["NoKeyInputTime"].Value = (seconds / 60.0).ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);

        ResidentTest.AppExitEndEnabled = chkAppExitEnd.Checked;
        config.AppSettings.Settings["AppExitEnd"].Value = ResidentTest.AppExitEndEnabled ? "1" : "0";

        ResidentTest.darkModeEnabled = chkDarkMode.Checked;
        config.AppSettings.Settings["IsDarkMode"].Value = ResidentTest.darkModeEnabled ? "on" : "off";

        // MS-IME入力設定(設定ファイルに項目が無い場合は追加する)
        SetAppSetting(config, "KeyAssignment", chkKeyAssign.Checked ? "1" : "0");
        SetAppSetting(config, "MsImeSwitch", chkMsImeSwitch.Checked ? "1" : "0");
        SetAppSetting(config, "SpaceWidth", $"{cmbSpaceInit.SelectedIndex},{cmbSpaceTo.SelectedIndex}");
        SetAppSetting(config, "Punctuation", $"{cmbPunctInit.SelectedIndex},{cmbPunctTo.SelectedIndex}");

        // 自動バックアップの設定
        if (cmbBackupHours.SelectedIndex >= 0)
            SetAppSetting(config, "BackupStableHours", backupHoursValues[cmbBackupHours.SelectedIndex].ToString("0.####", CultureInfo.InvariantCulture));
        if (cmbBackupKeep.SelectedIndex >= 0)
            SetAppSetting(config, "BackupKeep", backupKeepValues[cmbBackupKeep.SelectedIndex].ToString(CultureInfo.InvariantCulture));

        // 拡張設定で変更(リセット含む)したキーを書き込む。他の画面と同じキーを両方で変更した場合は拡張設定を優先する
        extGrid.EndEdit();
        foreach (var kv in extPending)
            SetAppSetting(config, kv.Key, kv.Value);

        config.Save(ConfigurationSaveMode.Modified);
        ConfigurationManager.RefreshSection("appSettings");

        // レジストリとトレイメニューへ即時反映する
        ResidentTest.ApplyMsImeSettings();

        // 拡張設定で他の画面と共通のキーを変更した場合に備えて、各画面を保存後の内容で更新する
        bool extChanged = extPending.Count > 0;
        if (extChanged)
            LoadCurrentValues();
        LoadExtendedSettings();
        if (extChanged)
            RenderCategory(selectedCategoryIndex);

        Debug.WriteLine("設定を保存しました");
        return true;
    }
}
