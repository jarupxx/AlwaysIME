using Microsoft.Win32;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;

// 設定ファイル（AlwaysIME.dll.config）の保存場所を管理する
//  ポータブル  : 実行ファイルと同じフォルダー
//  ユーザー別  : %AppData%\AlwaysIME
// 実行フォルダーに設定ファイルがあればポータブル、無くてユーザー別にあればユーザー別とする。
// どちらにも無い初回起動時だけ、保存場所を確認する。
static class ConfigStore
{
    public const string FileName = "AlwaysIME.dll.config";
    // 配布物に含める初期設定のひな形（無ければ既定値から生成する）
    public const string TemplateFileName = "AlwaysIME.dll.config.default";
    private const string UserFolderName = "AlwaysIME";

    public static string FilePath { get; private set; }
    public static bool IsPortable { get; private set; }

    private static string PortableDirectory => AppContext.BaseDirectory;
    private static string UserDirectory =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), UserFolderName);

    // 保存場所を決定して、ConfigurationManagerがそのファイルを読むようにする。
    // ConfigurationManagerを最初に使う前に呼ぶこと。falseの場合はアプリを終了する。
    public static bool Initialize()
    {
        string portablePath = Path.Combine(PortableDirectory, FileName);
        string userPath = Path.Combine(UserDirectory, FileName);

        if (File.Exists(portablePath))
        {
            FilePath = portablePath;
            IsPortable = true;
        }
        else if (File.Exists(userPath))
        {
            FilePath = userPath;
            IsPortable = false;
        }
        else
        {
            bool? portable = AskStorageMode();
            if (portable == null)
            {
                return false;
            }
            IsPortable = portable.Value;
            FilePath = IsPortable ? portablePath : userPath;
            try
            {
                CreateInitialConfigFile(FilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"設定ファイルを作成できませんでした。\n\n場所: {FilePath}\n{ex.Message}", "AlwaysIME", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        // System.Configurationが読み書きするファイルを切り替える
        AppDomain.CurrentDomain.SetData("APP_CONFIG_FILE", FilePath);
        return VerifyConfigurationManagerPath();
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
            MessageBox.Show($"設定ファイルの場所を切り替えられませんでした。\n\n想定: {FilePath}\n実際: {actual}", "AlwaysIME", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"設定ファイルを開けませんでした。\n\n場所: {FilePath}\n{ex.Message}", "AlwaysIME", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        return false;
    }

    // true: ポータブル / false: ユーザー別 / null: キャンセル
    private static bool? AskStorageMode()
    {
        bool canWritePortable = CanWriteDirectory(PortableDirectory);

        var userButton = new TaskDialogCommandLinkButton(
            "ユーザーごとに作成する（推奨）",
            "設定を「" + UserDirectory + "」に保存します。\nアプリを更新しても設定は消えません。");
        var portableButton = new TaskDialogCommandLinkButton(
            "ポータブルとして動作する",
            canWritePortable
                ? "設定を実行ファイルと同じフォルダーに保存します。\nフォルダーごと持ち運べます。更新時は設定ファイルを上書きしないでください。"
                : "このフォルダーには書き込めないため、選択できません。")
        {
            Enabled = canWritePortable,
        };

        var page = new TaskDialogPage
        {
            Caption = "AlwaysIME",
            Heading = "設定の保存場所を選んでください",
            Text = "初回起動のため、設定ファイルの保存場所を選びます。",
            Icon = TaskDialogIcon.Information,
            AllowCancel = true,
            DefaultButton = userButton,
            Buttons = { userButton, portableButton, TaskDialogButton.Cancel },
        };

        TaskDialogButton result = TaskDialog.ShowDialog(page);
        if (result == userButton)
        {
            return false;
        }
        if (result == portableButton)
        {
            return true;
        }
        return null;
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
            }
            catch (Exception ex)
            {
                Trace.WriteLine("ひな形を読み込めませんでした: " + ex.Message);
            }
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
