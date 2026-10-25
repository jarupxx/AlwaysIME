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
//
// 既定値（リセットの戻し先を含む）は、ひな形ファイル（AlwaysIME.dll.config.default）だけが持つ。
//  ・起動時にひな形を読み込み、ConfigSchema のルールで全キーを検証する。
//    無い・読めない・キーが不足・値が不正のいずれかなら、アプリは起動しない（ひな形の異常には関与しない）。
//  ・検証済みのひな形の値は Defaults としてメモリに保持し、以降はファイルを読み直さない。
//  ・利用者の設定ファイルの不足キーの補完と不正値の置き換えは RepairUserConfig()、
//    リセットは TryResetToTemplate() が、Defaults を使って行う。
static class ConfigStore
{
    public const string FileName = "AlwaysIME.dll.config";
    // 配布物に含める初期設定のひな形（必須。既定値の唯一の定義）
    public const string TemplateFileName = "AlwaysIME.dll.config.default";
    // 初回起動時に、このファイルが実行ファイルと同じフォルダーにあれば、ポータブルとして設定ファイルを作る（中身は問わない）
    public const string PortableMarkerFileName = "AlwaysIME.portable";
    private const string UserFolderName = "AlwaysIME";
    private const string LogFolderName = "logs";

    public static string FilePath { get; private set; }
    public static bool IsPortable { get; private set; }

    private static IReadOnlyDictionary<string, string> templateValues;

    // 検証済みのひな形の値（キー → 値）。Initialize() が成功した後でのみ使える
    public static IReadOnlyDictionary<string, string> Defaults =>
        templateValues ?? throw new InvalidOperationException("ひな形が読み込まれていません。");

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

    // 保存場所を決定し、ひな形を検証して、ConfigurationManagerがそのファイルを読むようにする。
    // ConfigurationManagerを最初に使う前に呼ぶこと。falseの場合はアプリを終了する。
    // （利用者の設定ファイルの補完・不正値の置き換えは、壊れたファイルの復旧（BackupService.OnStartup）の後に
    //   RepairUserConfig() で行う）
    public static bool Initialize()
    {
        // 1. 保存場所を決める（ファイルの有無だけで決まる）
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

        // 利用者への通知（書き込めない場所では、設定の保存・ログ・バックアップが働かないため、黙って動かさない）
        if (firstRun && markerExists && !portableWritable)
        {
            // マーカーがあるのに、ポータブルにできなかった
            MessageBox.Show(
                $"マーカーファイル({PortableMarkerFileName})がありますが、このフォルダーには書き込めません。\n\n" +
                $"フォルダー: {PortableDirectory}\n\n" +
                $"設定は、ユーザー別の場所（{UserDirectory}）に保存します。\n" +
                "ポータブルとして使うには、書き込める場所（例: デスクトップやドキュメント）へ移動してください。",
                "AlwaysIME", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        else if (portableExists && !CanWriteFile(portablePath))
        {
            // ポータブルの設定ファイルがあるのに、書き込めない場所で動いている（Program Files配下など）
            AppLog.Warn("設定ファイルに書き込めません（権限、または読み取り専用）: " + portablePath);
            MessageBox.Show(
                "設定ファイルに書き込めないため、設定の変更・ログ・バックアップは保存されない可能性があります。\n\n" +
                $"ファイル: {portablePath}\n\n" +
                "設定を保存するには、書き込める場所へ移動するか、設定ファイル(" + FileName + ")を削除して\n" +
                "ユーザー別の保存先(" + UserDirectory + ")で使ってください。",
                "AlwaysIME", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        // System.Configurationが読み書きするファイルを切り替える（以降のエラー表示でログ設定を読めるよう、先に行う）
        AppDomain.CurrentDomain.SetData("APP_CONFIG_FILE", FilePath);

        // 2. ひな形を読み込んで検証する。失敗したら起動しない
        string templatePath = Path.Combine(PortableDirectory, TemplateFileName);
        if (!TryLoadTemplate(templatePath, out Dictionary<string, string> loaded, out string kind, out string keys, out string detail, out Exception loadError))
        {
            return FailTemplate(kind, templatePath, keys, detail, loadError);
        }
        templateValues = loaded;
        AppLog.Info("ひな形を検証しました: " + templatePath);

        // 3. 初回起動なら、検証済みのひな形から設定ファイルを作る
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

        return VerifyConfigurationManagerPath();
    }

    // ひな形の異常を、画面とログに出す。ログの保存先は、設定ファイルの場所（決定済み）に合わせて確定する。
    private static bool FailTemplate(string kind, string templatePath, string keys, string detail, Exception ex)
    {
        AppLog.Error($"ひな形ファイルに問題があるため、起動を中止します（{kind}）: {templatePath}\n{detail}", ex);
        AppLog.Configure(IsPortable);

        var sb = new System.Text.StringBuilder();
        sb.AppendLine("設定ファイルに問題が見つかりました。\n対処方法がわからない場合は、PCの管理者や、PCの設定に詳しい方にご相談ください。");
        sb.AppendLine();
        sb.AppendLine($"ひな形ファイルに問題があります（{kind}）。");
        sb.AppendLine($"ファイル: {templatePath}");
        if (!string.IsNullOrEmpty(keys))
        {
            sb.AppendLine($"対象のキー: {keys}");
        }
        if (AppLog.IsEnabled)
        {
            sb.AppendLine();
            sb.Append($"詳細はログを参照してください: {AppLog.FilePath}");
        }
        MessageBox.Show(sb.ToString().TrimEnd(), "AlwaysIME", MessageBoxButtons.OK, MessageBoxIcon.Error);
        return false;
    }

    // ひな形を読み込み、ConfigSchema で全キーを検証する。ファイルの読み込みは1回だけで、結果は values に返す。
    //  kind   : 画面に出す原因の種別
    //  keys   : 画面に出す対象のキー（数が多い場合は省略）
    //  detail : ログに出す詳細（キー名、値、理由）
    internal static bool TryLoadTemplate(string path, out Dictionary<string, string> values,
        out string kind, out string keys, out string detail, out Exception error)
    {
        values = null;
        kind = keys = detail = null;
        error = null;

        if (!File.Exists(path))
        {
            kind = "ファイルが見つかりません";
            detail = "ひな形が存在しません。配布物の展開ミスや、ウイルス対策ソフトによる隔離などを確認してください。";
            return false;
        }

        XDocument doc;
        try
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            doc = XDocument.Load(fs);
        }
        catch (Exception ex)
        {
            kind = "読み込めません";
            detail = "ひな形をXMLとして読み込めませんでした: " + ex.Message;
            error = ex;
            return false;
        }

        XElement appSettings = doc.Root != null && doc.Root.Name.LocalName == "configuration"
            ? doc.Root.Element("appSettings")
            : null;
        if (appSettings == null)
        {
            kind = "形式が不正です";
            detail = "<configuration> の下に <appSettings> がありません。";
            return false;
        }

        var found = new Dictionary<string, string>(StringComparer.Ordinal);
        var problems = new List<string>();
        var problemKeys = new List<string>();
        var unknown = new List<string>();
        foreach (XElement add in appSettings.Elements("add"))
        {
            string key = (string)add.Attribute("key");
            if (string.IsNullOrEmpty(key))
            {
                problems.Add("  [不正] key属性がない <add> があります");
                continue;
            }
            if (found.ContainsKey(key))
            {
                problems.Add($"  [重複] {key}");
                problemKeys.Add(key);
                continue;
            }
            found[key] = (string)add.Attribute("value") ?? "";
            if (!ConfigSchema.IsKnownKey(key))
            {
                unknown.Add(key);
            }
        }

        bool anyMissing = false;
        bool anyInvalid = false;
        foreach (string key in ConfigSchema.Keys)
        {
            if (!found.TryGetValue(key, out string value))
            {
                problems.Add($"  [不足] {key}");
                problemKeys.Add(key);
                anyMissing = true;
                continue;
            }
            string reason = ConfigSchema.Check(key, value);
            if (reason != null)
            {
                problems.Add($"  [不正] {key} = \"{value}\" : {reason}");
                problemKeys.Add(key);
                anyInvalid = true;
            }
        }

        if (unknown.Count > 0)
        {
            AppLog.Warn("ひな形に未知のキーがあります（無視します）: " + string.Join(", ", unknown));
        }

        if (problems.Count > 0)
        {
            kind = anyMissing && anyInvalid ? "キーの不足と値の不正"
                : anyMissing ? "キーが不足しています"
                : anyInvalid ? "値が不正です"
                : "形式が不正です";
            const int shown = 5;
            keys = string.Join(", ", problemKeys.Take(shown)) + (problemKeys.Count > shown ? $" ほか{problemKeys.Count - shown}件" : "");
            detail = string.Join("\n", problems);
            return false;
        }

        values = found.Where(kv => ConfigSchema.IsKnownKey(kv.Key)).ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.Ordinal);
        return true;
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

    // 既存のファイルを書き込み用に開けるかを確認する（中身は書かない。更新日時も変わらない）
    private static bool CanWriteFile(string path)
    {
        try
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete);
            return true;
        }
        catch
        {
            return false;
        }
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

    // 検証済みのひな形の値で、設定ファイルを新規に作る（既存のファイルは置き換える）。
    // 初回起動と、壊れた設定ファイルを作り直すとき（BackupService）に使う。
    internal static void CreateInitialConfigFile(string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path));

        var appSettings = new XElement("appSettings");
        foreach (string key in ConfigSchema.Keys)
        {
            appSettings.Add(new XElement("add", new XAttribute("key", key), new XAttribute("value", Defaults[key])));
        }
        var doc = new XDocument(new XDeclaration("1.0", "utf-8", null), new XElement("configuration", appSettings));
        SaveAtomic(doc, path);
    }

    internal sealed class RepairResult
    {
        // ファイルに無かったため、ひな形の値で補ったキー
        public readonly List<string> Added = new List<string>();
        // 値が不正だったため、ひな形の値に置き換えたキー（元の値付き）
        public readonly List<(string Key, string OldValue, string NewValue, string Reason)> Replaced = new List<(string, string, string, string)>();
        // 修正した内容をファイルへ書き込めなかった
        public bool WriteFailed;
        public bool Changed => Added.Count > 0 || Replaced.Count > 0;
    }

    // 利用者の設定ファイルを検証して修復する。
    //  ・不足しているキーは、ひな形の値で補う
    //  ・不正な値は、そのキーだけひな形の値で置き換える（元の値はログに残す）
    //  ・ひな形に無いキーには触れない
    // 検証は ConfigurationManager ではなく XDocument で直接行う。
    // XMLとして読めないファイルはここでは扱わない（BackupService.OnStartup が復旧する）。
    // 1つ目のインスタンスだけが呼ぶこと。
    internal static RepairResult RepairUserConfig()
    {
        var result = new RepairResult();
        XDocument doc;
        try
        {
            doc = LoadConfigDocument(FilePath);
        }
        catch (Exception ex)
        {
            AppLog.Warn("設定ファイルを検証できませんでした: " + FilePath, ex);
            return result;
        }

        XElement appSettings = doc.Root.Element("appSettings");
        if (appSettings == null)
        {
            appSettings = new XElement("appSettings");
            doc.Root.Add(appSettings);
        }

        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (XElement add in appSettings.Elements("add").ToList())
        {
            string key = (string)add.Attribute("key");
            if (key == null || !ConfigSchema.IsKnownKey(key))
            {
                continue;
            }
            seen.Add(key);
            string value = (string)add.Attribute("value") ?? "";
            string reason = ConfigSchema.Check(key, value);
            if (reason != null)
            {
                string fixedValue = Defaults[key];
                add.SetAttributeValue("value", fixedValue);
                result.Replaced.Add((key, value, fixedValue, reason));
                AppLog.Warn($"設定値が不正だったため、ひな形の値に置き換えました: {key} 元の値=\"{value}\" → \"{fixedValue}\"（{reason}）");
            }
        }
        foreach (string key in ConfigSchema.Keys)
        {
            if (!seen.Contains(key))
            {
                string value = Defaults[key];
                appSettings.Add(new XElement("add", new XAttribute("key", key), new XAttribute("value", value)));
                result.Added.Add(key);
                AppLog.Info($"設定ファイルに無いキーを、ひな形の値で補いました: {key} = \"{value}\"");
            }
        }

        if (result.Changed)
        {
            try
            {
                SaveAtomic(doc, FilePath);
            }
            catch (Exception ex)
            {
                // 書き込めなくても起動は続ける（読み込み側が、不足・不正なキーをひな形の値で扱う）
                result.WriteFailed = true;
                AppLog.Warn("修復した設定を書き込めませんでした: " + FilePath, ex);
            }
        }
        return result;
    }

    // 設定を、ひな形の値に戻す（ひな形にある全キーが対象。ひな形に無いキーはそのまま残す）。
    // 実行中も呼べる。ひな形は起動時に検証済みのものを使うので、ファイルは読み直さない。
    // 成功したら、呼び出し側が ConfigurationManager.RefreshSection("appSettings") と
    // ResidentTest.ApplyMsImeSettings() で反映すること（ファイルへの書き込みが成功してから反映する）。
    internal static bool TryResetToTemplate(out string error)
    {
        error = null;
        if (templateValues == null)
        {
            error = "ひな形が読み込まれていないため、リセットできません。";
            return false;
        }

        XDocument doc;
        try
        {
            doc = LoadConfigDocument(FilePath);
        }
        catch (Exception ex)
        {
            AppLog.Error("リセットを中止しました。設定ファイルを読み込めません: " + FilePath, ex);
            error = "設定ファイルを読み込めないため、リセットできません。\n" + ex.Message;
            return false;
        }

        // 誤操作から戻せるよう、リセット前の設定を保存する（失敗してもリセット自体は続ける）
        try
        {
            if (!BackupService.BackupNow(out string backupMessage))
            {
                AppLog.Info("リセット前のバックアップ: " + backupMessage);
            }
        }
        catch (Exception ex)
        {
            AppLog.Warn("リセット前のバックアップに失敗しました", ex);
        }

        try
        {
            XElement appSettings = doc.Root.Element("appSettings");
            if (appSettings == null)
            {
                appSettings = new XElement("appSettings");
                doc.Root.Add(appSettings);
            }
            foreach (string key in ConfigSchema.Keys)
            {
                string value = Defaults[key];
                var targets = appSettings.Elements("add").Where(e => (string)e.Attribute("key") == key).ToList();
                if (targets.Count == 0)
                {
                    appSettings.Add(new XElement("add", new XAttribute("key", key), new XAttribute("value", value)));
                }
                foreach (XElement target in targets)
                {
                    target.SetAttributeValue("value", value);
                }
            }
            SaveAtomic(doc, FilePath);
            AppLog.Info("設定をひな形の値にリセットしました: " + FilePath);
            return true;
        }
        catch (Exception ex)
        {
            AppLog.Error("設定をリセットできませんでした: " + FilePath, ex);
            error = "設定ファイルに書き込めませんでした。\n" + ex.Message;
            return false;
        }
    }

    // 設定ファイルをXDocumentとして読む（ConfigurationManagerを経由しない）。XMLとして読めなければ例外
    private static XDocument LoadConfigDocument(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        XDocument doc = XDocument.Load(fs);
        if (doc.Root == null || doc.Root.Name.LocalName != "configuration")
        {
            throw new InvalidDataException("ルート要素が <configuration> ではありません。");
        }
        return doc;
    }

    // 書き込み途中で止まっても不完全なファイルが残らないよう、一時ファイル経由で置き換える
    private static void SaveAtomic(XDocument doc, string path)
    {
        string tempPath = path + ".tmp";
        try
        {
            doc.Save(tempPath);
            File.Move(tempPath, path, overwrite: true);
        }
        catch
        {
            try { File.Delete(tempPath); } catch { }
            throw;
        }
    }
}
