using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

// Windowsへのログオン時にAlwaysIMEを自動で起動する仕組み。
//  方法1: スタートアップフォルダーにショートカットを作成する
//  方法2: タスクスケジューラに「ログオン時」のタスクを登録する（ログオンの10秒後に起動）
//  方法3: 方法2のタスクを「最上位の特権で実行する」で登録する（管理者権限で動くウィンドウを操作するため）
// どれも現在のユーザーだけが対象。方法1と方法2は、管理者権限は不要。
// 方法3の登録・変更・解除は、AlwaysIMEを管理者として起動しているときだけできる（通常権限からは、登録内容を読むだけ）。
// 2つの方法は同時に使わない
// （2つ目のインスタンスは再起動シグナルを送るため、二重に起動すると動作が乱れる）。
//
// 登録の有無は、設定ファイルではなく実際のショートカット／タスクで判断する。
// 設定ファイルは複数のユーザーで共有されることがある（ポータブル）ため。
enum AutoStartMode
{
    None,
    StartupFolder,
    TaskScheduler,
    TaskSchedulerElevated,   // タスクスケジューラ（最上位の特権で実行する）
}

sealed class AutoStartState
{
    public bool ShortcutExists;
    public bool? TaskExists;       // nullは、タスクスケジューラの状態を確認できなかった場合
    public bool? TaskElevated;     // タスクが「最上位の特権」で登録されているか。nullは、確認できなかった場合
    public bool Conflict => ShortcutExists && TaskExists == true;
    public AutoStartMode Mode =>
        TaskExists == true ? (TaskElevated == true ? AutoStartMode.TaskSchedulerElevated : AutoStartMode.TaskScheduler)
        : ShortcutExists ? AutoStartMode.StartupFolder
        : AutoStartMode.None;

    // 管理者権限で登録されていて、いまのプロセスが通常権限のため、変更できない
    public bool LockedForCurrentProcess =>
        Mode == AutoStartMode.TaskSchedulerElevated && !AppLog.IsElevatedProcess;
}

static class AutoStart
{
    private const string ShortcutFileName = "AlwaysIME.lnk";
    private const string TaskFolder = @"\AlwaysIME\";
    private const string TaskDelay = "PT10S";        // ログオンの10秒後
    private const int ProcessTimeoutMs = 20000;

    // 登録した内容の控え（現在のユーザーのレジストリ）。
    // 実行ファイルを移動したときに、起動時に登録を更新するために使う。
    private const string HintKeyPath = @"Software\AlwaysIME";
    private const string HintModeValue = "AutoStartMode";
    private const string HintPathValue = "AutoStartPath";

    [DllImport("kernel32.dll")]
    private static extern uint GetOEMCP();

    private static string ShortcutPath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Startup), ShortcutFileName);

    // ---------------------------------------------------------------
    // 現在の状態
    // ---------------------------------------------------------------
    public static AutoStartState Detect()
    {
        var state = new AutoStartState();
        try
        {
            state.ShortcutExists = File.Exists(ShortcutPath);
        }
        catch (Exception ex)
        {
            AppLog.Warn("スタートアップのショートカットを確認できませんでした", ex);
        }

        // /XML で登録内容を取得する（通常権限でも読める）。RunLevel が HighestAvailable なら、最上位の特権で登録されている
        int code = RunSchtasks($"/Query /TN \"{GetTaskName()}\" /XML", out string output);
        if (code == 0)
        {
            state.TaskExists = true;
            Match level = Regex.Match(output, @"<RunLevel>\s*(\w+)\s*</RunLevel>");
            state.TaskElevated = level.Success
                ? string.Equals(level.Groups[1].Value, "HighestAvailable", StringComparison.OrdinalIgnoreCase)
                : (bool?)null;
        }
        else if (code > 0)
        {
            state.TaskExists = false;
        }
        else
        {
            // schtasks.exeを実行できなかった（制限されている場合など）
            AppLog.Warn("タスクスケジューラの状態を確認できませんでした: " + output);
            state.TaskExists = null;
        }
        return state;
    }

    // ---------------------------------------------------------------
    // 登録・解除
    // ---------------------------------------------------------------

    // 選んだ方法で登録し、もう一方の登録は削除する。失敗したときは、理由をerrorに返す。
    public static bool TryApply(AutoStartMode mode, out string error)
    {
        error = null;
        var problems = new List<string>();
        try
        {
            // 管理者権限の登録は、管理者権限のプロセスだけが変更できる
            if (!AppLog.IsElevatedProcess)
            {
                if (mode == AutoStartMode.TaskSchedulerElevated)
                {
                    throw new InvalidOperationException("管理者権限での自動起動は、AlwaysIMEを管理者として起動しているときだけ登録できます。");
                }
                if (Detect().Mode == AutoStartMode.TaskSchedulerElevated)
                {
                    throw new InvalidOperationException("管理者権限で登録された自動起動は、AlwaysIMEを管理者として起動して変更してください。");
                }
            }

            string exe = mode == AutoStartMode.None ? null : GetExePath();
            bool isTask = mode == AutoStartMode.TaskScheduler || mode == AutoStartMode.TaskSchedulerElevated;

            // 失敗したときに今の登録を失わないよう、選んだ方法を先に登録する
            if (mode == AutoStartMode.StartupFolder)
            {
                CreateShortcut(exe);
            }
            else if (isTask)
            {
                CreateTask(exe, mode == AutoStartMode.TaskSchedulerElevated);
            }
            SaveHint(mode, exe);

            if (mode != AutoStartMode.StartupFolder)
            {
                TryRemoveShortcut(problems);
            }
            if (!isTask)
            {
                TryRemoveTask(problems);
            }

            AppLog.Info("自動起動の設定を変更しました: " + Describe(mode));
        }
        catch (Exception ex)
        {
            AppLog.Warn("自動起動の設定を変更できませんでした: " + Describe(mode), ex);
            error = ex.Message;
            return false;
        }

        if (problems.Count > 0)
        {
            error = string.Join("\n", problems);
            AppLog.Warn("自動起動の設定の一部を変更できませんでした: " + error);
            return false;
        }
        return true;
    }

    // 起動時に呼ぶ。実行ファイルを移動していたら、登録内容を新しい場所に更新する。
    // 通常（場所が変わっていない）は、レジストリを読むだけで何もしない。
    public static void VerifyRegisteredPath()
    {
        try
        {
            string modeText;
            string savedPath;
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(HintKeyPath))
            {
                modeText = key?.GetValue(HintModeValue) as string;
                savedPath = key?.GetValue(HintPathValue) as string;
            }

            AutoStartMode mode = modeText == "task" ? AutoStartMode.TaskScheduler
                : modeText == "task-admin" ? AutoStartMode.TaskSchedulerElevated
                : modeText == "startup" ? AutoStartMode.StartupFolder
                : AutoStartMode.None;
            if (mode == AutoStartMode.None)
            {
                return;
            }
            // 管理者権限の登録は、通常権限のプロセスからは更新できない
            if (mode == AutoStartMode.TaskSchedulerElevated && !AppLog.IsElevatedProcess)
            {
                return;
            }

            string current = GetExePath();
            if (string.Equals(savedPath, current, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            // 利用者が手動で解除していた場合に、勝手に登録し直さないよう、登録が残っているときだけ更新する
            bool stillRegistered = mode == AutoStartMode.StartupFolder
                ? File.Exists(ShortcutPath)
                : RunSchtasks($"/Query /TN \"{GetTaskName()}\"", out _) == 0;
            if (!stillRegistered)
            {
                return;
            }

            AppLog.Info($"実行ファイルの場所が変わったため、自動起動の登録を更新します: {savedPath} → {current}");
            if (!TryApply(mode, out string error))
            {
                AppLog.Warn("自動起動の登録を更新できませんでした: " + error);
            }
        }
        catch (Exception ex)
        {
            AppLog.Warn("自動起動の登録を確認できませんでした", ex);
        }
    }

    private static string Describe(AutoStartMode mode)
    {
        switch (mode)
        {
            case AutoStartMode.StartupFolder: return "スタートアップのショートカット";
            case AutoStartMode.TaskScheduler: return "タスクスケジューラ";
            case AutoStartMode.TaskSchedulerElevated: return "タスクスケジューラ（管理者権限）";
            default: return "自動起動しない";
        }
    }

    // 実行ファイルが、管理者しか書き換えられない場所（Program Files、Windows）にあるか。
    // 最上位の特権で自動起動する場合、ユーザー権限で書き換えられる場所の実行ファイルは、差し替えによって管理者権限を得られてしまう。
    internal static bool IsProtectedInstallLocation()
    {
        try
        {
            string dir = Path.GetFullPath(Path.GetDirectoryName(GetExePath()) ?? "");
            var roots = new[]
            {
                Environment.SpecialFolder.ProgramFiles,
                Environment.SpecialFolder.ProgramFilesX86,
                Environment.SpecialFolder.Windows,
            };
            foreach (var folder in roots)
            {
                string root = Environment.GetFolderPath(folder);
                if (root.Length > 0 && dir.StartsWith(root.TrimEnd('\\') + "\\", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
        }
        catch
        {
            // 判定できなければ、保護されていないものとして扱う
        }
        return false;
    }

    private static string GetExePath()
    {
        string path = Environment.ProcessPath;
        if (string.IsNullOrEmpty(path))
        {
            throw new InvalidOperationException("実行ファイルの場所を取得できませんでした。");
        }
        if (Path.GetFileNameWithoutExtension(path).Equals("dotnet", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("dotnetコマンド経由で実行されているため、登録できません。AlwaysIME.exe を直接起動してください。");
        }
        return path;
    }

    // ---------------------------------------------------------------
    // スタートアップのショートカット
    // ---------------------------------------------------------------
    private static void CreateShortcut(string exePath)
    {
        string path = ShortcutPath;
        Directory.CreateDirectory(Path.GetDirectoryName(path));

        var link = (IShellLinkW)new ShellLink();
        try
        {
            link.SetPath(exePath);
            link.SetWorkingDirectory(Path.GetDirectoryName(exePath));
            link.SetDescription("AlwaysIME");
            link.SetIconLocation(exePath, 0);
            ((IPersistFile)link).Save(path, false);
        }
        finally
        {
            Marshal.FinalReleaseComObject(link);
        }
    }

    private static void TryRemoveShortcut(List<string> problems)
    {
        try
        {
            string path = ShortcutPath;
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (Exception ex)
        {
            problems.Add("スタートアップのショートカットを削除できませんでした: " + ex.Message);
        }
    }

    [ComImport, Guid("00021401-0000-0000-C000-000000000046")]
    private class ShellLink
    {
    }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214F9-0000-0000-C000-000000000046")]
    private interface IShellLinkW
    {
        void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cch, IntPtr pfd, int fFlags);
        void GetIDList(out IntPtr ppidl);
        void SetIDList(IntPtr pidl);
        void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cch);
        void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cch);
        void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
        void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cch);
        void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
        void GetHotkey(out short pwHotkey);
        void SetHotkey(short wHotkey);
        void GetShowCmd(out int piShowCmd);
        void SetShowCmd(int iShowCmd);
        void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cch, out int piIcon);
        void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
        void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, int dwReserved);
        void Resolve(IntPtr hwnd, int fFlags);
        void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
    }

    // ---------------------------------------------------------------
    // タスクスケジューラ（schtasks.exe にXMLを渡して登録する）
    // ---------------------------------------------------------------

    // タスクはPC全体で名前が共有されるため、ユーザーごとに別の名前にする
    internal static string MakeTaskName(string userName, string sid)
    {
        string user = Regex.Replace(string.IsNullOrEmpty(userName) ? "user" : userName, @"[^A-Za-z0-9_.-]", "_");
        string rid = "";
        if (!string.IsNullOrEmpty(sid) && sid.Contains('-'))
        {
            rid = "_" + sid.Substring(sid.LastIndexOf('-') + 1);
        }
        return TaskFolder + "AutoStart_" + user + rid;
    }

    private static string GetTaskName()
    {
        string sid = null;
        try
        {
            using (WindowsIdentity id = WindowsIdentity.GetCurrent())
            {
                sid = id.User?.Value;
            }
        }
        catch
        {
            // SIDを取得できなくても、ユーザー名だけで名前を作る
        }
        return MakeTaskName(Environment.UserName, sid);
    }

    // ログオン時に起動するタスクのXML。elevated が true なら「最上位の特権で実行する」、false なら現在のユーザー権限（昇格なし）
    internal static string BuildTaskXml(string user, string exePath, bool elevated = false)
    {
        XNamespace ns = "http://schemas.microsoft.com/windows/2004/02/mit/task";
        var task = new XElement(ns + "Task",
            new XAttribute("version", "1.2"),
            new XElement(ns + "RegistrationInfo",
                new XElement(ns + "Description", "AlwaysIME をログオン時に起動します。")),
            new XElement(ns + "Triggers",
                new XElement(ns + "LogonTrigger",
                    new XElement(ns + "Enabled", "true"),
                    new XElement(ns + "UserId", user),
                    new XElement(ns + "Delay", TaskDelay))),
            new XElement(ns + "Principals",
                new XElement(ns + "Principal",
                    new XAttribute("id", "Author"),
                    new XElement(ns + "UserId", user),
                    new XElement(ns + "LogonType", "InteractiveToken"),
                    new XElement(ns + "RunLevel", elevated ? "HighestAvailable" : "LeastPrivilege"))),
            new XElement(ns + "Settings",
                new XElement(ns + "MultipleInstancesPolicy", "IgnoreNew"),
                new XElement(ns + "DisallowStartIfOnBatteries", "false"),
                new XElement(ns + "StopIfGoingOnBatteries", "false"),
                new XElement(ns + "AllowHardTerminate", "true"),
                new XElement(ns + "StartWhenAvailable", "true"),
                new XElement(ns + "RunOnlyIfNetworkAvailable", "false"),
                new XElement(ns + "IdleSettings",
                    new XElement(ns + "StopOnIdleEnd", "false"),
                    new XElement(ns + "RestartOnIdle", "false")),
                new XElement(ns + "AllowStartOnDemand", "true"),
                new XElement(ns + "Enabled", "true"),
                new XElement(ns + "Hidden", "false"),
                new XElement(ns + "RunOnlyIfIdle", "false"),
                new XElement(ns + "WakeToRun", "false"),
                new XElement(ns + "ExecutionTimeLimit", "PT0S"),
                new XElement(ns + "Priority", "7")),
            new XElement(ns + "Actions",
                new XAttribute("Context", "Author"),
                new XElement(ns + "Exec",
                    new XElement(ns + "Command", exePath),
                    new XElement(ns + "WorkingDirectory", Path.GetDirectoryName(exePath)))));

        return "<?xml version=\"1.0\" encoding=\"UTF-16\"?>\r\n" + task.ToString();
    }

    private static void CreateTask(string exePath, bool elevated)
    {
        string user;
        using (WindowsIdentity id = WindowsIdentity.GetCurrent())
        {
            user = id.Name;
        }

        string xmlPath = Path.Combine(Path.GetTempPath(), "AlwaysIME_task_" + Guid.NewGuid().ToString("N") + ".xml");
        try
        {
            // schtasks.exe はUTF-16のXMLを受け付ける
            File.WriteAllText(xmlPath, BuildTaskXml(user, exePath, elevated), Encoding.Unicode);
            int code = RunSchtasks($"/Create /TN \"{GetTaskName()}\" /XML \"{xmlPath}\" /F", out string output);
            if (code != 0)
            {
                throw new InvalidOperationException("タスクスケジューラへ登録できませんでした。" + (output.Length > 0 ? "\n" + output : ""));
            }
        }
        finally
        {
            try
            {
                File.Delete(xmlPath);
            }
            catch
            {
                // 一時ファイルを消せなくても影響はない
            }
        }
    }

    private static void TryRemoveTask(List<string> problems)
    {
        string name = GetTaskName();
        int exists = RunSchtasks($"/Query /TN \"{name}\"", out _);
        if (exists != 0)
        {
            // 登録されていない（または確認できない）場合は、削除するものがない
            return;
        }
        int code = RunSchtasks($"/Delete /TN \"{name}\" /F", out string output);
        if (code != 0)
        {
            problems.Add("タスクスケジューラの登録を削除できませんでした: " + output);
        }
    }

    // schtasks.exe を実行する。終了コードを返す（起動できなかった場合や時間切れは -1）。
    private static int RunSchtasks(string arguments, out string output)
    {
        output = "";
        try
        {
            // PATHに依存せず、システムフォルダーのものを実行する
            string exe = Path.Combine(Environment.SystemDirectory, "schtasks.exe");
            Encoding oem = GetOemEncoding();
            var psi = new ProcessStartInfo(exe, arguments)
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = oem,
                StandardErrorEncoding = oem,
            };
            using (Process p = Process.Start(psi))
            {
                if (p == null)
                {
                    output = "schtasks.exe を起動できませんでした。";
                    return -1;
                }
                // 出力が溜まって固まらないよう、先に読み取りを始める
                Task<string> stdout = p.StandardOutput.ReadToEndAsync();
                Task<string> stderr = p.StandardError.ReadToEndAsync();
                if (!p.WaitForExit(ProcessTimeoutMs))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch
                    {
                        // すでに終了している場合など
                    }
                    output = "schtasks.exe が時間内に終了しませんでした。";
                    return -1;
                }
                output = (stdout.Result + stderr.Result).Trim();
                return p.ExitCode;
            }
        }
        catch (Exception ex)
        {
            output = ex.Message;
            return -1;
        }
    }

    // schtasks.exe は、コンソールのコードページ（OEM）で出力する
    private static Encoding GetOemEncoding()
    {
        try
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            return Encoding.GetEncoding((int)GetOEMCP());
        }
        catch
        {
            return Encoding.UTF8;
        }
    }

    // ---------------------------------------------------------------
    // 登録内容の控え
    // ---------------------------------------------------------------
    private static void SaveHint(AutoStartMode mode, string exePath)
    {
        try
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(HintKeyPath))
            {
                if (mode == AutoStartMode.None)
                {
                    key.DeleteValue(HintModeValue, false);
                    key.DeleteValue(HintPathValue, false);
                }
                else
                {
                    key.SetValue(HintModeValue,
                        mode == AutoStartMode.TaskScheduler ? "task"
                        : mode == AutoStartMode.TaskSchedulerElevated ? "task-admin"
                        : "startup");
                    key.SetValue(HintPathValue, exePath);
                }
            }
        }
        catch (Exception ex)
        {
            // 控えを保存できなくても、自動起動自体は動作する（移動時の自動更新ができなくなるだけ）
            AppLog.Warn("自動起動の登録内容を保存できませんでした", ex);
        }
    }
}
