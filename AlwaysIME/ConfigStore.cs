using System.Configuration;
using System.Xml.Linq;

// 設定ファイル（AlwaysIME.dll.config）の保存場所を管理する
//  ポータブル  : 実行ファイルと同じフォルダー
//  ユーザー別  : %AppData%\AlwaysIME
// 決め方（利用者への確認はしない）:
//  1. 実行フォルダーに設定ファイルがあれば、ポータブル
//  2. 無くて、ユーザー別の場所にあれば、ユーザー別
//  3. どちらにも無い初回起動では、実行フォルダーにマーカーファイル（AlwaysIME.portable）があればポータブル、
//     無ければユーザー別で作る
static class ConfigStore
{
    public const string FileName = "AlwaysIME.dll.config";
    // 配布物に含める初期設定のひな形（無ければ既定値から生成する）
    public const string TemplateFileName = "AlwaysIME.dll.config.default";
    // 初回起動時に、このファイルが実行ファイルと同じフォルダーにあれば、ポータブルとして設定ファイルを作る（中身は問わない）
    public const string PortableMarkerFileName = "AlwaysIME.portable";
    private const string UserFolderName = "AlwaysIME";
    private const string LogFolderName = "logs";

    public static string FilePath { get; private set; }
    public static bool IsPortable { get; private set; }

    private static string PortableDirectory => AppContext.BaseDirectory;
    private static string UserDirectory =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), UserFolderName);

    public static bool HasPortableMarker =>
        File.Exists(Path.Combine(PortableDirectory, PortableMarkerFileName));

    // ログの保存先。設定ファイルと同じ方式で決まる。
    //  ポータブル  : 実行ファイルと同じフォルダー\logs
    //  ユーザー別  : %AppData%\AlwaysIME\logs
    public static string GetLogDirectory(bool portable) =>
        Path.Combine(portable ? PortableDirectory : UserDirectory, LogFolderName);

    // 既存の設定ファイルの場所を、確認ダイアログなしで調べる。ポータブル=true、ユーザー別=false、どちらも無い=null。
    public static bool? PeekIsPortable()
    {
        if (File.Exists(Path.Combine(PortableDirectory, FileName)))
        {
            return true;
        }
        if (File.Exists(Path.Combine(UserDirectory, FileName)))
        {
            return false;
        }
        return null;
    }

    // 既存の設定ファイルがあれば、ConfigurationManagerがそれを読むようにする（確認ダイアログは出さない）。
    // 2つ目のインスタンスが、ログの設定を読むために使う。ConfigurationManagerを最初に使う前に呼ぶこと。
    public static bool TryActivateExisting()
    {
        bool? portable = PeekIsPortable();
        if (portable == null)
        {
            return false;
        }
        IsPortable = portable.Value;
        FilePath = Path.Combine(IsPortable ? PortableDirectory : UserDirectory, FileName);
        AppDomain.CurrentDomain.SetData("APP_CONFIG_FILE", FilePath);
        return true;
    }

    // 保存場所を決定して、ConfigurationManagerがそのファイルを読むようにする。
    // ConfigurationManagerを最初に使う前に呼ぶこと。falseの場合はアプリを終了する。
    public static bool Initialize()
    {
        string portablePath = Path.Combine(PortableDirectory, FileName);
        string userPath = Path.Combine(UserDirectory, FileName);
        bool portableExists = File.Exists(portablePath);
        bool userExists = File.Exists(userPath);
        bool firstRun = !portableExists && !userExists;
        bool markerExists = HasPortableMarker;
        // 書き込めるかの確認（一時ファイルを作る）は、マーカーで保存場所を決める初回だけ行う
        bool portableWritable = firstRun && markerExists && CanWriteDirectory(PortableDirectory);

        var decision = DecideLocation(portableExists, userExists, markerExists, portableWritable);
        IsPortable = decision.Portable;
        FilePath = IsPortable ? portablePath : userPath;
        if (decision.Warn)
        {
            AppLog.Warn(decision.Reason);
        }
        else
        {
            AppLog.Info(decision.Reason);
        }

        if (firstRun)
        {
            try
            {
                CreateInitialConfigFile(FilePath);
            }
            catch (Exception ex)
            {
                AppLog.Error("設定ファイルを作成できませんでした: " + FilePath, ex);
                MessageBox.Show($"設定ファイルを作成できませんでした。\n\n場所: {FilePath}\n{ex.Message}", "AlwaysIME", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // System.Configurationが読み書きするファイルを切り替える
        AppDomain.CurrentDomain.SetData("APP_CONFIG_FILE", FilePath);
        return VerifyConfigurationManagerPath();
    }

    // 設定ファイルの保存場所を決める（ファイルの有無の組み合わせだけで決まる。副作用なし）。
    //  ・既存の設定ファイルがあれば、それを使う（実行フォルダー → ユーザー別の順）。マーカーより優先する
    //  ・初回起動（どちらにも無い）ときだけ、マーカーファイルでポータブルにする
    //  ・マーカーがあっても実行フォルダーに書き込めないときは、ユーザー別にして警告を残す
    internal static (bool Portable, string Reason, bool Warn) DecideLocation(
        bool portableConfigExists, bool userConfigExists, bool markerExists, bool portableWritable)
    {
        if (portableConfigExists)
        {
            return (true, "実行ファイルのフォルダーに設定ファイルがあるため、ポータブルとして動作します", false);
        }
        if (userConfigExists)
        {
            return markerExists
                ? (false, "ユーザー別の設定ファイルが既にあるため、マーカーファイル(" + PortableMarkerFileName + ")は使わず、ユーザー別として動作します", true)
                : (false, "ユーザー別の設定ファイルがあるため、ユーザー別として動作します", false);
        }
        if (markerExists)
        {
            return portableWritable
                ? (true, "初回起動: マーカーファイルがあるため、ポータブルとして設定ファイルを作成します", false)
                : (false, "初回起動: マーカーファイルがありますが、実行ファイルのフォルダーに書き込めないため、ユーザー別として設定ファイルを作成します", true);
        }
        return (false, "初回起動: ユーザー別として設定ファイルを作成します", false);
    }

    // 切り替えが効いていないと、意図しない場所へ書き込む恐れがあるため起動時に確認する
    private static bool VerifyConfigurationManagerPath()
    {
        try
        {
            string actual = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None).FilePath;
            if (string.Equals(Path.GetFullPath(actual), Path.GetFullPath(FilePath), StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            AppLog.Error($"設定ファイルの場所を切り替えられませんでした。想定: {FilePath} / 実際: {actual}");
            MessageBox.Show($"設定ファイルの場所を切り替えられませんでした。\n\n想定: {FilePath}\n実際: {actual}", "AlwaysIME", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        catch (Exception ex)
        {
            AppLog.Error("設定ファイルを開けませんでした: " + FilePath, ex);
            MessageBox.Show($"設定ファイルを開けませんでした。\n\n場所: {FilePath}\n{ex.Message}", "AlwaysIME", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        return false;
    }

    private static bool CanWriteDirectory(string dir)
    {
        try
        {
            string testPath = Path.Combine(dir, ".write-test-" + Guid.NewGuid().ToString("N"));
            File.WriteAllText(testPath, "");
            File.Delete(testPath);
            return true;
        }
        catch
        {
            return false;
        }
    }

    // ひな形（AlwaysIME.dll.config.default）をもとに設定ファイルを作る。
    // ひな形に無いキーは既定値で補い、全てのキーが必ず存在する状態にする。
    internal static void CreateInitialConfigFile(string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path));

        XDocument doc = null;
        string templatePath = Path.Combine(PortableDirectory, TemplateFileName);
        if (File.Exists(templatePath))
        {
            try
            {
                doc = XDocument.Load(templatePath);
                AppLog.Info("ひな形をもとに設定ファイルを作成します: " + templatePath);
            }
            catch (Exception ex)
            {
                AppLog.Warn("ひな形を読み込めませんでした", ex);
            }
        }
        if (doc == null)
        {
            AppLog.Info("ひな形を使わず、既定値で設定ファイルを作成します（ひな形の場所: " + templatePath + "）");
        }
        doc ??= new XDocument(new XDeclaration("1.0", "utf-8", null));

        XElement root = doc.Root;
        if (root == null)
        {
            root = new XElement("configuration");
            doc.Add(root);
        }
        XElement appSettings = root.Element("appSettings");
        if (appSettings == null)
        {
            appSettings = new XElement("appSettings");
            root.Add(appSettings);
        }

        var existingKeys = new HashSet<string>(
            appSettings.Elements("add").Select(e => (string)e.Attribute("key")).Where(k => k != null));
        foreach (var kv in ResidentTest.DefaultConfigValues)
        {
            if (!existingKeys.Contains(kv.Key))
            {
                appSettings.Add(new XElement("add", new XAttribute("key", kv.Key), new XAttribute("value", kv.Value)));
            }
        }

        // 書き込み途中で止まっても不完全なファイルが残らないよう、一時ファイル経由で置き換える
        string tempPath = path + ".tmp";
        doc.Save(tempPath);
        File.Move(tempPath, path, overwrite: true);
    }
}
