using System.Configuration;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

// ファイルへのログ出力。
//  場所          : ポータブル  = 実行ファイルと同じフォルダー\logs\AlwaysIME.log
//                  ユーザー別  = %AppData%\AlwaysIME\logs\AlwaysIME.log
//                  （設定ファイルの保存場所と同じ方式。決まるまでのログはメモリに溜めて、決まったら書き出す）
//  ローテーション : 1ファイルが最大サイズを超えたら AlwaysIME.1.log → .2.log … へ繰り下げ、超えた世代は削除する
//  設定          : AlwaysIME.dll.config の LogEnabled / LogLevel / LogMaxSizeMB / LogKeep
//                  （設定画面の「上級者向け設定」から変更できる。ReloadSettings()で反映する）
//  記録しないもの : ウィンドウタイトル、入力内容など、利用者の作業内容が分かる情報
//                   （そのような情報は Debug.WriteLine を使うこと。Releaseビルドでは出力されない）
//  安全性        : ログの書き込みに失敗しても例外は外へ出さず、アプリの動作を止めない。
//  重複の抑制    : 同じ内容のメッセージは60秒に1回だけ書き、省略した件数を次回の行に付ける。
//                  （毎周期の監視処理から繰り返し出るメッセージで、ログが溢れるのを防ぐ）
static class AppLog
{
    private const string BaseName = "AlwaysIME";

    // 設定値の範囲と既定値（設定画面と共通）
    public const string DefaultLevel = "info";
    public const int DefaultMaxSizeMB = 1;
    public const int MinMaxSizeMB = 1;
    public const int MaxMaxSizeMB = 100;
    public const int DefaultKeep = 3;
    public const int MinKeep = 1;
    public const int MaxKeep = 20;

    private const int SeverityInfo = 0;
    private const int SeverityWarn = 1;
    private const int SeverityError = 2;
    private const int SeverityFatal = 3;

    private const int MaxPending = 200;          // 保存先が決まるまでに溜めておく最大件数
    private const int MaxTrackedMessages = 500;  // 重複判定のために覚えておくメッセージの最大数
    private static readonly TimeSpan DuplicateWindow = TimeSpan.FromSeconds(60);
    private static readonly UTF8Encoding Utf8WithBom = new UTF8Encoding(true);

    private sealed class DuplicateState
    {
        public DateTime LastWritten;
        public int Suppressed;
    }

    private static readonly object gate = new object();
    private static readonly Dictionary<string, DuplicateState> duplicates = new Dictionary<string, DuplicateState>();
    private static readonly List<(int Severity, string Text)> pending = new List<(int Severity, string Text)>();

    private static bool locationDecided;
    private static long currentSize;

    // 現在の設定（保存先が決まる前は既定値）
    private static bool enabled = true;
    private static int minSeverity = SeverityInfo;
    private static long maxFileBytes = DefaultMaxSizeMB * 1024L * 1024L;
    private static int keepGenerations = DefaultKeep;
    private static string appliedSummary;

    // ログファイルのパス。保存先が決まる前、または保存先を用意できなかった場合はnull
    public static string FilePath { get; private set; }
    // ログフォルダーのパス（同上）
    public static string LogDirectory { get; private set; }

    // 実際にファイルへ書き込む状態かどうか
    public static bool IsEnabled
    {
        get
        {
            lock (gate)
            {
                return enabled && FilePath != null;
            }
        }
    }

    // ---------------------------------------------------------------
    // 設定値の解釈（設定画面からも使う。不正な値は既定値にする）
    // ---------------------------------------------------------------
    public static bool ParseEnabled(string s)
    {
        if (string.IsNullOrWhiteSpace(s))
        {
            return true;
        }
        string v = s.Trim();
        return !(v == "0" || v.Equals("off", StringComparison.OrdinalIgnoreCase) || v.Equals("false", StringComparison.OrdinalIgnoreCase));
    }

    // "info" / "warn" / "error" のいずれかを返す
    public static string ParseLevel(string s)
    {
        string v = (s ?? "").Trim().ToLowerInvariant();
        return v == "warn" || v == "error" ? v : DefaultLevel;
    }

    public static int ParseMaxSizeMB(string s)
    {
        return int.TryParse(s, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int v)
            ? Math.Clamp(v, MinMaxSizeMB, MaxMaxSizeMB)
            : DefaultMaxSizeMB;
    }

    public static int ParseKeep(string s)
    {
        return int.TryParse(s, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int v)
            ? Math.Clamp(v, MinKeep, MaxKeep)
            : DefaultKeep;
    }

    private static int LevelToSeverity(string level)
    {
        return level == "error" ? SeverityError : level == "warn" ? SeverityWarn : SeverityInfo;
    }

    // ---------------------------------------------------------------
    // 保存先と設定
    // ---------------------------------------------------------------

    // 設定ファイルの保存場所が決まったら呼ぶ。保存先を確定し、それまでに溜めたログを書き出す。
    // ConfigurationManagerが設定ファイルを読める状態で呼ぶこと。
    public static void Configure(bool portable)
    {
        try
        {
            ReadSettings(out bool newEnabled, out string newLevel, out int newSizeMB, out int newKeep);
            lock (gate)
            {
                if (locationDecided)
                {
                    return;
                }

                string failNote = null;
                string dir = TryPrepareDirectory(ConfigStore.GetLogDirectory(portable), out Exception firstError);
                if (dir == null && portable)
                {
                    // 実行ファイルのフォルダーが書き込み不可の場合は、ユーザー別の場所に切り替える
                    failNote = "実行ファイルのフォルダーにログを作成できないため、ユーザー別の場所を使用します: " + firstError?.Message;
                    dir = TryPrepareDirectory(ConfigStore.GetLogDirectory(false), out _);
                }

                locationDecided = true;
                if (dir != null)
                {
                    LogDirectory = dir;
                    FilePath = Path.Combine(dir, BaseName + ".log");
                    currentSize = File.Exists(FilePath) ? new FileInfo(FilePath).Length : 0;
                }

                ApplySettingsLocked(newEnabled, newLevel, newSizeMB, newKeep, announce: false);
                FlushPendingLocked();
                if (failNote != null)
                {
                    WriteLocked(SeverityWarn, failNote, null);
                }
                if (enabled && FilePath != null)
                {
                    WriteLocked(SeverityInfo, "ログ設定: " + appliedSummary, null, force: true);
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine("ログの初期化に失敗しました: " + ex.Message);
        }
    }

    // 設定ファイルを読める状態になる前に終了する場合などに、既存の設定ファイルの場所からログの保存先を決める
    public static void EnsureLocation()
    {
        try
        {
            lock (gate)
            {
                if (locationDecided)
                {
                    return;
                }
            }
            Configure(ConfigStore.PeekIsPortable() ?? ConfigStore.HasPortableMarker);
        }
        catch
        {
            // ログの準備に失敗してもアプリを止めない
        }
    }

    // 設定ファイルの変更後（読み込み時・保存時）に呼ぶ。
    public static void ReloadSettings()
    {
        try
        {
            ReadSettings(out bool newEnabled, out string newLevel, out int newSizeMB, out int newKeep);
            lock (gate)
            {
                ApplySettingsLocked(newEnabled, newLevel, newSizeMB, newKeep, announce: true);
            }
        }
        catch
        {
            // 設定の反映に失敗しても、現在の設定のまま動作を続ける
        }
    }

    private static void ReadSettings(out bool newEnabled, out string newLevel, out int newSizeMB, out int newKeep)
    {
        try
        {
            newEnabled = ParseEnabled(ConfigurationManager.AppSettings["LogEnabled"]);
            newLevel = ParseLevel(ConfigurationManager.AppSettings["LogLevel"]);
            newSizeMB = ParseMaxSizeMB(ConfigurationManager.AppSettings["LogMaxSizeMB"]);
            newKeep = ParseKeep(ConfigurationManager.AppSettings["LogKeep"]);
        }
        catch
        {
            // 設定ファイルが読めない場合は既定値で動作する
            newEnabled = true;
            newLevel = DefaultLevel;
            newSizeMB = DefaultMaxSizeMB;
            newKeep = DefaultKeep;
        }
    }

    private static void ApplySettingsLocked(bool newEnabled, string newLevel, int newSizeMB, int newKeep, bool announce)
    {
        string summary = $"出力={(newEnabled ? "あり" : "なし")}, レベル={newLevel}, 最大サイズ={newSizeMB}MB, 過去ログ={newKeep}個";
        bool canAnnounce = announce && locationDecided && FilePath != null;

        if (canAnnounce && enabled && !newEnabled)
        {
            // 無効にする直前に、その旨を記録する
            WriteLocked(SeverityInfo, "ログ出力を無効にします（設定変更）", null, force: true);
        }

        enabled = newEnabled;
        minSeverity = LevelToSeverity(newLevel);
        maxFileBytes = newSizeMB * 1024L * 1024L;
        keepGenerations = newKeep;

        if (locationDecided && FilePath != null)
        {
            PruneExtraGenerations();
        }
        if (canAnnounce && newEnabled && summary != appliedSummary)
        {
            WriteLocked(SeverityInfo, "ログ設定: " + summary, null, force: true);
        }
        appliedSummary = summary;
    }

    private static string TryPrepareDirectory(string dir, out Exception error)
    {
        error = null;
        try
        {
            Directory.CreateDirectory(dir);
            return dir;
        }
        catch (Exception ex)
        {
            error = ex;
            return null;
        }
    }

    // ---------------------------------------------------------------
    // 書き込み
    // ---------------------------------------------------------------

    // 起動時の環境情報を記録する（プライバシーに関わる情報は含めない）
    public static void LogStartup()
    {
        Info($"==== AlwaysIME 起動 v{GetVersion()} ====");
        Info($"OS: {RuntimeInformation.OSDescription} / .NET: {RuntimeInformation.FrameworkDescription} / {(Environment.Is64BitProcess ? "64bit" : "32bit")}");
        Info($"PID: {Environment.ProcessId} / セッションID: {GetSessionId()} / 管理者権限: {IsElevated()}");
        Info($"実行ファイル: {Environment.ProcessPath}");
    }

    public static void Info(string message) => Write(SeverityInfo, message, null);
    public static void Warn(string message, Exception ex = null) => Write(SeverityWarn, message, ex);
    public static void Error(string message, Exception ex = null) => Write(SeverityError, message, ex);
    public static void Fatal(string message, Exception ex = null) => Write(SeverityFatal, message, ex);

    private static void Write(int severity, string message, Exception ex)
    {
        // デバッガー接続時に見られるよう、Debugビルドでは出力ウィンドウにも出す
        Debug.WriteLine($"[{SeverityLabel(severity)}] {message}");

        try
        {
            lock (gate)
            {
                WriteLocked(severity, message, ex);
            }
        }
        catch
        {
            // ログが書けないことでアプリを止めない
        }
    }

    private static bool ShouldWrite(int severity)
    {
        return enabled && severity >= minSeverity;
    }

    // force=true: ログのレベル設定に関わらず記録する（設定変更の記録など）。ただしログ出力が無効の間は書かない。
    private static void WriteLocked(int severity, string message, Exception ex, bool force = false)
    {
        if (locationDecided && (FilePath == null || !(force ? enabled : ShouldWrite(severity))))
        {
            return;
        }

        DateTime now = DateTime.Now;
        if (IsDuplicate(severity, message, ex, now, out string note))
        {
            return;
        }

        string text = Format(now, severity, message, ex, note);
        if (!locationDecided)
        {
            // 保存先が決まるまでは溜めておく（出力するかどうかは、決まった後の設定で判断する）
            if (pending.Count >= MaxPending)
            {
                pending.RemoveAt(0);
            }
            pending.Add((severity, text));
            return;
        }
        AppendText(text);
    }

    private static void FlushPendingLocked()
    {
        foreach (var item in pending)
        {
            if (FilePath != null && ShouldWrite(item.Severity))
            {
                try
                {
                    AppendText(item.Text);
                }
                catch
                {
                    // 1件の失敗で残りを捨てない
                }
            }
        }
        pending.Clear();
    }

    private static string Format(DateTime now, int severity, string message, Exception ex, string note)
    {
        var sb = new StringBuilder();
        sb.Append(now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
        sb.Append(" [").Append(SeverityLabel(severity)).Append("] [T").Append(Environment.CurrentManagedThreadId).Append("] ");
        sb.Append(message);
        if (note != null)
        {
            sb.Append(' ').Append(note);
        }
        if (ex != null)
        {
            // 例外の型、メッセージ、スタックトレース、InnerExceptionをまとめて出力する
            sb.AppendLine();
            sb.Append(ex);
        }
        sb.AppendLine();
        return sb.ToString();
    }

    private static string SeverityLabel(int severity)
    {
        switch (severity)
        {
            case SeverityWarn: return "WRN";
            case SeverityError: return "ERR";
            case SeverityFatal: return "FTL";
            default: return "INF";
        }
    }

    private static void AppendText(string text)
    {
        int bytes = Utf8WithBom.GetByteCount(text);
        RotateIfNeeded(bytes);

        // 実行中に利用者やサポート担当者がログを開けるよう、共有を許可して開く
        using (var fs = new FileStream(FilePath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
        using (var sw = new StreamWriter(fs, Utf8WithBom))
        {
            sw.Write(text);
        }
        currentSize += bytes;
    }

    // 直近に同じメッセージを書いていれば true（今回は書かない）。
    // 省略した件数があれば、今回の行に付ける注記を note に返す。
    private static bool IsDuplicate(int severity, string message, Exception ex, DateTime now, out string note)
    {
        note = null;
        string key = severity + "|" + message + "|" + (ex == null ? "" : ex.GetType().FullName + ":" + ex.Message);

        if (duplicates.TryGetValue(key, out DuplicateState state))
        {
            if (now - state.LastWritten < DuplicateWindow)
            {
                state.Suppressed++;
                return true;
            }
            if (state.Suppressed > 0)
            {
                note = $"（直近{(int)DuplicateWindow.TotalSeconds}秒間に同じメッセージを{state.Suppressed}件省略）";
                state.Suppressed = 0;
            }
            state.LastWritten = now;
            return false;
        }

        if (duplicates.Count >= MaxTrackedMessages)
        {
            duplicates.Clear();
        }
        duplicates[key] = new DuplicateState { LastWritten = now };
        return false;
    }

    // ---------------------------------------------------------------
    // ローテーション
    // ---------------------------------------------------------------
    private static string GenerationPath(int generation) =>
        Path.Combine(LogDirectory, $"{BaseName}.{generation}.log");

    private static void RotateIfNeeded(int incomingBytes)
    {
        if (currentSize + incomingBytes <= maxFileBytes)
        {
            return;
        }

        try
        {
            // AlwaysIME.2.log → .3.log、.1.log → .2.log、AlwaysIME.log → .1.log の順に繰り下げる
            for (int i = keepGenerations; i >= 1; i--)
            {
                string src = i == 1 ? FilePath : GenerationPath(i - 1);
                string dst = GenerationPath(i);
                if (!File.Exists(src))
                {
                    continue;
                }
                File.Move(src, dst, true);
            }
            PruneExtraGenerations();
        }
        catch
        {
            // 他のプログラムがログを開いている場合など。書き込み自体は続ける。
        }
        // 失敗した場合に、書き込みのたびに繰り下げを試みないよう、サイズは常にリセットする
        currentSize = 0;
    }

    // 世代数を減らした設定に変えた場合などに、範囲外の古い世代を削除する
    private static void PruneExtraGenerations()
    {
        try
        {
            for (int i = keepGenerations + 1; i <= MaxKeep; i++)
            {
                string path = GenerationPath(i);
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
        }
        catch
        {
            // 削除できなくても動作には影響しない
        }
    }

    // ---------------------------------------------------------------
    // 環境情報
    // ---------------------------------------------------------------
    private static string GetVersion()
    {
        try
        {
            return Assembly.GetEntryAssembly()?.GetName().Version?.ToString() ?? "不明";
        }
        catch
        {
            return "不明";
        }
    }

    private static string GetSessionId()
    {
        try
        {
            using (Process p = Process.GetCurrentProcess())
            {
                return p.SessionId.ToString();
            }
        }
        catch
        {
            return "不明";
        }
    }

    private static string IsElevated()
    {
        try
        {
            using (var identity = System.Security.Principal.WindowsIdentity.GetCurrent())
            {
                var principal = new System.Security.Principal.WindowsPrincipal(identity);
                return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator) ? "あり" : "なし";
            }
        }
        catch
        {
            return "不明";
        }
    }
}

// 未処理の例外をログに残す。
//  ・UIスレッド（タイマー、メニュー操作など）の例外 : ログに記録して動作を継続する
//  ・UIスレッド以外の例外（プロセスが終了する）   : ログに記録し、利用者にも終了を知らせる
//  ・監視されなかったTaskの例外                    : ログに記録する
static class GlobalExceptionHandler
{
    private static int fatalDialogShown;

    // フォームやコントロールを作る前に呼ぶこと
    public static void Install()
    {
        try
        {
            // 設定ファイルの指定に関わらず、UIスレッドの例外を ThreadException で受ける
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        }
        catch (Exception ex)
        {
            AppLog.Warn("未処理例外モードを設定できませんでした", ex);
        }
        Application.ThreadException += OnThreadException;
        AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
    }

    private static void OnThreadException(object sender, ThreadExceptionEventArgs e)
    {
        AppLog.EnsureLocation();
        AppLog.Error("UIスレッドで未処理の例外が発生しました（動作を継続します）", e.Exception);
    }

    private static void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
    {
        AppLog.EnsureLocation();
        Exception ex = e.ExceptionObject as Exception;
        string detail = ex == null ? " 例外オブジェクト: " + e.ExceptionObject : "";
        AppLog.Fatal($"未処理の例外が発生しました（プロセス終了: {e.IsTerminating}）{detail}", ex);

        if (e.IsTerminating && System.Threading.Interlocked.Exchange(ref fatalDialogShown, 1) == 0)
        {
            try
            {
                string where = AppLog.IsEnabled ? AppLog.FilePath : "（ログ出力が無効のため、記録されていません）";
                MessageBox.Show(
                    "AlwaysIMEで予期しないエラーが発生したため、終了します。\n\n"
                    + "ログ: " + where,
                    "AlwaysIME", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch
            {
                // 終了処理中なので、表示に失敗しても何もしない
            }
        }
    }

    private static void OnUnobservedTaskException(object sender, UnobservedTaskExceptionEventArgs e)
    {
        AppLog.Warn("監視されなかったTaskの例外が発生しました", e.Exception);
        e.SetObserved();
    }
}
