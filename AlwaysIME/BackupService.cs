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

// 設定ファイルのバックアップを管理する
//   <設定ファイルの場所>\backup\
//     latest.config                   : 最後に確認できた有効な設定。終了時と変更検出時に更新する（常に最新）
//     stable-<連番>-<UTC日時>.config  : 変更後に一定時間動作を続けた＝安定動作を確認できた設定。複数世代を保持する
//     backup.json                     : 連番と安定判定のための状態
// ・新旧の判定は連番で行い、ファイルの日時は使わない（システム時刻を変更しても影響を受けない）。
// ・安定の判定は時計ではなく、アプリの累積稼働時間で行う。設定を変更するとリセットされるため、
//   試行錯誤している間は作成されず、変更後に落ち着いたものだけが残る。
// ・正常に終了しなかった場合も、安定の確認は最初からやり直す。
static class BackupService
{
    private const string BackupFolderName = "backup";
    private const string LatestFileName = "latest.config";
    private const string StateFileName = "backup.json";
    private const string LegacyBackupFileName = "~AlwaysIME.dll.config.bak";
    private const double DefaultStableHours = 24;
    private const int DefaultKeep = 5;
    private const int TickMinutes = 10;
    // スリープ中などで経過時間が飛んだときに、稼働として数えすぎないための上限
    private const double MaxAccumulateSeconds = TickMinutes * 60 * 2;
    private const int DebounceMilliseconds = 2000;
    private const int MaxRestoreCandidates = 4;

    private static readonly Regex StableNameRegex = new Regex(@"^stable-(\d+)-(\d{14})\.config$", RegexOptions.Compiled);
    private static readonly Regex PreRestoreNameRegex = new Regex(@"^pre-restore-(\d+)-(\d{14})\.config$", RegexOptions.Compiled);
    private const int PreRestoreKeep = 3;
    private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions { WriteIndented = true };

    public sealed class BackupState
    {
        public int Version { get; set; } = 1;
        // 最後に割り当てた安定版の連番
        public long StableSeq { get; set; }
        // 最後に確認した設定ファイルのハッシュ
        public string ConfigHash { get; set; } = "";
        // ConfigHash の設定で動作を続けた累積時間（秒）
        public double StableRunSeconds { get; set; }
        // 前回、正常に終了したか（初回は true）
        public bool CleanExit { get; set; } = true;
        // 最後に割り当てた「復元前の設定」の連番
        public long PreRestoreSeq { get; set; }
    }

    public enum BackupKind
    {
        Latest,      // latest.config
        Stable,      // stable-*.config
        PreRestore,  // pre-restore-*.config（復元する直前の設定）
    }

    // 復元する範囲
    public enum RestoreScope
    {
        All,         // すべての設定
        ListsOnly,   // アプリ一覧・連携設定のみ
        ExceptLists, // それ以外の設定のみ
    }

    public sealed class BackupEntry
    {
        public string FilePath;
        public BackupKind Kind;
        public long Seq;          // Latest は 0
        public DateTime SavedUtc; // 表示用（新旧の判定には使わない）
    }

    // 「アプリ一覧・連携設定」として扱うキー
    private static readonly HashSet<string> ListSettingKeys = new HashSet<string>
    {
        "PassList", "ImeOffList", "ImeOffTitle", "UpdateTag",
        "OnActivatedAppList", "OnActivatedAppPath", "OnActivatedArgv",
        "BackgroundAppList", "BackgroundAppPath", "BackgroundArgv",
    };

    private static readonly object gate = new object();
    private static string configPath;
    private static string backupDir;
    private static BackupState state = new BackupState();
    private static long lastMark;
    private static bool started;
    private static bool exited;
    private static System.Threading.Timer tickTimer;
    private static System.Threading.Timer debounceTimer;
    private static FileSystemWatcher watcher;

    // 起動時に呼ぶ。falseの場合はアプリを終了する。
    // バックアップ機能自体が使えない場合でも、本体の動作は妨げない（trueを返す）。
    public static bool OnStartup()
    {
        lock (gate)
        {
            try
            {
                configPath = ConfigStore.FilePath;
                backupDir = Path.Combine(Path.GetDirectoryName(configPath), BackupFolderName);
                Directory.CreateDirectory(backupDir);
            }
            catch (Exception ex)
            {
                Trace.WriteLine("バックアップフォルダーを作成できませんでした: " + ex.Message);
                return true;
            }

            state = LoadState();
            ImportLegacyBackup();

            // 設定ファイルが壊れていたら、バックアップから復元する
            if (!TryReadConfig(configPath, out _, out _))
            {
                if (!RecoverBrokenConfig())
                {
                    return false;
                }
            }

            TryReadConfig(configPath, out _, out string hash);
            hash ??= "";
            bool unclean = !state.CleanExit;
            if (hash != state.ConfigHash || unclean)
            {
                // 設定が変わった、または前回が正常終了でなかった: 安定の確認をやり直す
                state.ConfigHash = hash;
                state.StableRunSeconds = 0;
            }
            state.CleanExit = false;

            SyncLatest();
            SaveState();

            lastMark = Stopwatch.GetTimestamp();
            started = true;
            StartMonitoring();
            return true;
        }
    }

    // 終了時に呼ぶ（何度呼んでも1回だけ実行する）
    public static void OnExit()
    {
        lock (gate)
        {
            if (!started || exited)
            {
                return;
            }
            exited = true;
            try
            {
                CheckConfigChange();
                TryPromoteToStable();
                state.CleanExit = true;
                SaveState();
            }
            catch (Exception ex)
            {
                Trace.WriteLine("終了時のバックアップに失敗しました: " + ex.Message);
            }
            try
            {
                watcher?.Dispose();
                tickTimer?.Dispose();
                debounceTimer?.Dispose();
            }
            catch { }
        }
    }

    // ---- 設定画面から使うAPI ----

    public static string BackupDirectory
    {
        get
        {
            lock (gate)
            {
                return backupDir;
            }
        }
    }

    // 復元に使えるバックアップの一覧。
    // 最新の設定 → 安定版（連番の新しい順） → 復元前の設定（連番の新しい順）
    public static List<BackupEntry> GetBackups()
    {
        lock (gate)
        {
            var list = new List<BackupEntry>();
            if (backupDir == null)
            {
                return list;
            }
            string latestPath = Path.Combine(backupDir, LatestFileName);
            if (File.Exists(latestPath))
            {
                list.Add(new BackupEntry { FilePath = latestPath, Kind = BackupKind.Latest, SavedUtc = File.GetLastWriteTimeUtc(latestPath) });
            }
            list.AddRange(ListStables());
            list.AddRange(ListPreRestores());
            return list;
        }
    }

    // 保存済みの設定を、安定の確認を待たずに安定版として保存する
    public static bool BackupNow(out string message)
    {
        lock (gate)
        {
            message = "バックアップ機能を利用できません。";
            if (!started || exited)
            {
                return false;
            }
            try
            {
                CheckConfigChange();
                bool created = CreateStableBackup(out message);
                if (created)
                {
                    state.CleanExit = false;
                    SaveState();
                }
                return created;
            }
            catch (Exception ex)
            {
                message = "バックアップを作成できませんでした: " + ex.Message;
                return false;
            }
        }
    }

    // バックアップから設定ファイルへ復元する。現在の設定は「復元前の設定」として残す。
    // バックアップに無いキーは現在の値を保つ（古いバックアップで新しい項目が消えないようにするため）。
    public static bool RestoreFrom(BackupEntry entry, RestoreScope scope, out string error)
    {
        lock (gate)
        {
            error = "バックアップ機能を利用できません。";
            if (!started || exited)
            {
                return false;
            }
            try
            {
                if (!TryReadConfig(entry.FilePath, out byte[] backupBytes, out _))
                {
                    error = "選択したバックアップを読み込めません。";
                    return false;
                }
                if (!TryReadConfig(configPath, out byte[] currentBytes, out _))
                {
                    error = "現在の設定ファイルを読み込めません。";
                    return false;
                }
                XDocument backupDoc = XDocument.Load(new MemoryStream(backupBytes));
                XDocument currentDoc = XDocument.Load(new MemoryStream(currentBytes));
                XElement backupApp = backupDoc.Root.Element("appSettings");
                if (backupApp == null)
                {
                    error = "バックアップに設定が含まれていません。";
                    return false;
                }
                XElement currentApp = currentDoc.Root.Element("appSettings");
                if (currentApp == null)
                {
                    currentApp = new XElement("appSettings");
                    currentDoc.Root.Add(currentApp);
                }

                // 現在の設定を残してから書き換える（元に戻せるようにする）
                List<BackupEntry> preRestores = ListPreRestores();
                long preSeq = Math.Max(state.PreRestoreSeq, preRestores.Count > 0 ? preRestores[0].Seq : 0) + 1;
                WriteAtomic(Path.Combine(backupDir, $"pre-restore-{preSeq:D6}-{DateTime.UtcNow:yyyyMMddHHmmss}.config"), currentBytes);
                state.PreRestoreSeq = preSeq;
                List<BackupEntry> allPre = ListPreRestores();
                for (int i = PreRestoreKeep; i < allPre.Count; i++)
                {
                    try { File.Delete(allPre[i].FilePath); } catch { }
                }

                foreach (XElement add in backupApp.Elements("add"))
                {
                    string key = (string)add.Attribute("key");
                    if (key == null)
                    {
                        continue;
                    }
                    bool isList = ListSettingKeys.Contains(key);
                    if ((scope == RestoreScope.ListsOnly && !isList) || (scope == RestoreScope.ExceptLists && isList))
                    {
                        continue;
                    }
                    string value = (string)add.Attribute("value") ?? "";
                    XElement target = currentApp.Elements("add").FirstOrDefault(e => (string)e.Attribute("key") == key);
                    if (target != null)
                    {
                        target.SetAttributeValue("value", value);
                    }
                    else
                    {
                        currentApp.Add(new XElement("add", new XAttribute("key", key), new XAttribute("value", value)));
                    }
                }

                string tempPath = configPath + ".tmp";
                currentDoc.Save(tempPath);
                File.Move(tempPath, configPath, overwrite: true);

                // 変更を取り込み、安定の確認をやり直す
                CheckConfigChange();
                state.CleanExit = false;
                SaveState();
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }
    }

    private static void StartMonitoring()
    {
        AppDomain.CurrentDomain.ProcessExit += (s, e) => OnExit();
        SystemEvents.SessionEnding += (s, e) => OnExit();

        tickTimer = new System.Threading.Timer(_ => OnTick(), null,
            TimeSpan.FromMinutes(TickMinutes), TimeSpan.FromMinutes(TickMinutes));
        debounceTimer = new System.Threading.Timer(_ => OnConfigChanged(), null,
            Timeout.Infinite, Timeout.Infinite);

        try
        {
            // 設定ファイルの変更を検出する（設定画面・メニュー・テキストエディターでの編集のいずれも対象）
            watcher = new FileSystemWatcher(Path.GetDirectoryName(configPath), Path.GetFileName(configPath))
            {
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.Size | NotifyFilters.CreationTime,
            };
            FileSystemEventHandler handler = (s, e) => ScheduleConfigCheck();
            watcher.Changed += handler;
            watcher.Created += handler;
            watcher.Renamed += (s, e) => ScheduleConfigCheck();
            watcher.EnableRaisingEvents = true;
        }
        catch (Exception ex)
        {
            Trace.WriteLine("設定ファイルの監視を開始できませんでした: " + ex.Message);
        }
    }

    // 保存が連続しても1回にまとめる
    private static void ScheduleConfigCheck()
    {
        try
        {
            debounceTimer?.Change(DebounceMilliseconds, Timeout.Infinite);
        }
        catch (ObjectDisposedException) { }
    }

    private static void OnTick()
    {
        lock (gate)
        {
            if (!started || exited)
            {
                return;
            }
            try
            {
                CheckConfigChange();
                TryPromoteToStable();
                state.CleanExit = false;
                SaveState();
            }
            catch (Exception ex)
            {
                Trace.WriteLine("バックアップの定期確認に失敗しました: " + ex.Message);
            }
        }
    }

    private static void OnConfigChanged()
    {
        lock (gate)
        {
            if (!started || exited)
            {
                return;
            }
            try
            {
                CheckConfigChange();
                state.CleanExit = false;
                SaveState();
            }
            catch (Exception ex)
            {
                Trace.WriteLine("設定変更の確認に失敗しました: " + ex.Message);
            }
        }
    }

    // 稼働時間を加算し、設定が変わっていれば安定の確認をやり直して latest を更新する
    private static void CheckConfigChange()
    {
        AccumulateRunTime();

        // 編集途中などで読めない・壊れている間は何もしない（良いバックアップを上書きしない）
        if (!TryReadConfig(configPath, out byte[] bytes, out string hash))
        {
            return;
        }
        if (hash != state.ConfigHash)
        {
            state.ConfigHash = hash;
            state.StableRunSeconds = 0;
        }
        SyncLatest(bytes, hash);
    }

    private static void AccumulateRunTime()
    {
        long now = Stopwatch.GetTimestamp();
        double delta = Stopwatch.GetElapsedTime(lastMark, now).TotalSeconds;
        lastMark = now;
        if (delta < 0)
        {
            delta = 0;
        }
        state.StableRunSeconds += Math.Min(delta, MaxAccumulateSeconds);
    }

    private static void SyncLatest()
    {
        if (TryReadConfig(configPath, out byte[] bytes, out string hash))
        {
            SyncLatest(bytes, hash);
        }
    }

    private static void SyncLatest(byte[] bytes, string hash)
    {
        string latestPath = Path.Combine(backupDir, LatestFileName);
        try
        {
            if (TryReadConfig(latestPath, out _, out string latestHash) && latestHash == hash)
            {
                return;
            }
            WriteAtomic(latestPath, bytes);
        }
        catch (Exception ex)
        {
            Trace.WriteLine("latest の更新に失敗しました: " + ex.Message);
        }
    }

    // 変更後に十分な時間動作を続けていれば、安定版として保存する
    private static void TryPromoteToStable()
    {
        if (state.StableRunSeconds < GetStableHours() * 3600)
        {
            return;
        }
        CreateStableBackup(out _);
    }

    // 現在の設定を安定版として保存する。作成した場合は true（しなかった理由は result に入る）
    private static bool CreateStableBackup(out string result)
    {
        result = "";
        if (!TryReadConfig(configPath, out byte[] bytes, out string hash))
        {
            result = "設定ファイルを読み込めないため、作成できませんでした。";
            return false;
        }
        if (hash != state.ConfigHash)
        {
            result = "設定ファイルが変更されました。もう一度お試しください。";
            return false;
        }

        List<BackupEntry> stables = ListStables();
        // 直近の安定版と同じ内容なら作らない
        if (stables.Count > 0 && TryReadConfig(stables[0].FilePath, out _, out string newestHash) && newestHash == hash)
        {
            result = "直近の安定版と同じ内容のため、作成しませんでした。";
            return false;
        }

        long seq = Math.Max(state.StableSeq, stables.Count > 0 ? stables[0].Seq : 0) + 1;
        string name = $"stable-{seq:D6}-{DateTime.UtcNow:yyyyMMddHHmmss}.config";
        WriteAtomic(Path.Combine(backupDir, name), bytes);
        state.StableSeq = seq;
        Trace.WriteLine("安定版バックアップを作成しました: " + name);

        // 古い世代を削除する
        int keep = GetKeepCount();
        List<BackupEntry> all = ListStables();
        for (int i = keep; i < all.Count; i++)
        {
            try
            {
                File.Delete(all[i].FilePath);
            }
            catch (Exception ex)
            {
                Trace.WriteLine("古いバックアップを削除できませんでした: " + ex.Message);
            }
        }

        result = $"安定版 #{seq} を作成しました。";
        return true;
    }

    private static double GetStableHours()
    {
        try
        {
            string s = ConfigurationManager.AppSettings["BackupStableHours"];
            if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out double v)
                && v >= 0 && !double.IsInfinity(v))
            {
                return v;
            }
        }
        catch { }
        return DefaultStableHours;
    }

    private static int GetKeepCount()
    {
        try
        {
            string s = ConfigurationManager.AppSettings["BackupKeep"];
            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int v))
            {
                return Math.Clamp(v, 1, 50);
            }
        }
        catch { }
        return DefaultKeep;
    }

    // 安定版を連番の新しい順に返す
    private static List<BackupEntry> ListStables() =>
        ListBackupFiles("stable-*.config", StableNameRegex, BackupKind.Stable);

    // 復元前の設定を連番の新しい順に返す
    private static List<BackupEntry> ListPreRestores() =>
        ListBackupFiles("pre-restore-*.config", PreRestoreNameRegex, BackupKind.PreRestore);

    private static List<BackupEntry> ListBackupFiles(string pattern, Regex regex, BackupKind kind)
    {
        var list = new List<BackupEntry>();
        try
        {
            foreach (string file in Directory.EnumerateFiles(backupDir, pattern))
            {
                Match m = regex.Match(Path.GetFileName(file));
                if (!m.Success || !long.TryParse(m.Groups[1].Value, out long seq))
                {
                    continue;
                }
                DateTime saved = DateTime.TryParseExact(m.Groups[2].Value, "yyyyMMddHHmmss", CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out DateTime t)
                    ? t
                    : File.GetLastWriteTimeUtc(file);
                list.Add(new BackupEntry { FilePath = file, Kind = kind, Seq = seq, SavedUtc = saved });
            }
        }
        catch (Exception ex)
        {
            Trace.WriteLine("バックアップ一覧を取得できませんでした: " + ex.Message);
        }
        list.Sort((a, b) => b.Seq.CompareTo(a.Seq));
        return list;
    }

    // 旧方式（~AlwaysIME.dll.config.bak）のバックアップを、最初の安定版として引き継ぐ
    private static void ImportLegacyBackup()
    {
        try
        {
            string legacyPath = Path.Combine(Path.GetDirectoryName(configPath), LegacyBackupFileName);
            if (!File.Exists(legacyPath) || state.StableSeq != 0 || ListStables().Count > 0)
            {
                return;
            }
            if (!TryReadConfig(legacyPath, out byte[] bytes, out _))
            {
                return;
            }
            string name = $"stable-{1:D6}-{File.GetLastWriteTimeUtc(legacyPath):yyyyMMddHHmmss}.config";
            WriteAtomic(Path.Combine(backupDir, name), bytes);
            state.StableSeq = 1;
            File.Delete(legacyPath);
            Trace.WriteLine("旧バックアップを引き継ぎました: " + name);
        }
        catch (Exception ex)
        {
            Trace.WriteLine("旧バックアップの引き継ぎに失敗しました: " + ex.Message);
        }
    }

    // 設定ファイルが読めないときの復旧。壊れたファイルは corrupt-*.config として残す。
    // 戻り値 false: 復旧せずに終了する
    private static bool RecoverBrokenConfig()
    {
        var candidates = new List<BackupEntry>();
        string latestPath = Path.Combine(backupDir, LatestFileName);
        if (File.Exists(latestPath) && TryReadConfig(latestPath, out _, out _))
        {
            candidates.Add(new BackupEntry { FilePath = latestPath, Kind = BackupKind.Latest, SavedUtc = File.GetLastWriteTimeUtc(latestPath) });
        }
        foreach (BackupEntry stable in ListStables())
        {
            if (candidates.Count >= MaxRestoreCandidates)
            {
                break;
            }
            if (TryReadConfig(stable.FilePath, out _, out _))
            {
                candidates.Add(stable);
            }
        }

        var page = new TaskDialogPage
        {
            Caption = "AlwaysIME",
            Heading = "設定ファイルを読み込めません",
            Text = "設定ファイルが壊れている可能性があります。\n場所: " + configPath
                + "\n\n復元する内容を選んでください。壊れた設定ファイルは backup フォルダーに「corrupt-…」として残します。",
            Icon = TaskDialogIcon.Warning,
            AllowCancel = true,
        };
        var buttons = new List<(TaskDialogCommandLinkButton Button, BackupEntry Entry)>();
        foreach (BackupEntry entry in candidates)
        {
            string title = entry.Kind == BackupKind.Latest ? "最新のバックアップから復元" : $"安定版（連番 {entry.Seq}）から復元";
            string saved = $"保存: {entry.SavedUtc.ToLocalTime():yyyy/MM/dd HH:mm:ss}";
            var button = new TaskDialogCommandLinkButton(title, saved);
            page.Buttons.Add(button);
            buttons.Add((button, entry));
        }
        var freshButton = new TaskDialogCommandLinkButton("初期設定で作り直す", "全ての設定が初期状態になります。");
        page.Buttons.Add(freshButton);
        page.Buttons.Add(TaskDialogButton.Cancel);

        TaskDialogButton result = TaskDialog.ShowDialog(page);
        try
        {
            if (result == freshButton)
            {
                MoveCorruptedConfigAside();
                ConfigStore.CreateInitialConfigFile(configPath);
                return true;
            }
            foreach (var (button, entry) in buttons)
            {
                if (result == button)
                {
                    TryReadConfig(entry.FilePath, out byte[] bytes, out _);
                    MoveCorruptedConfigAside();
                    WriteAtomic(configPath, bytes);
                    return true;
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show("設定ファイルを復元できませんでした。\n\n" + ex.Message, "AlwaysIME", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        return false;
    }

    private static void MoveCorruptedConfigAside()
    {
        if (File.Exists(configPath))
        {
            string corruptPath = Path.Combine(backupDir, $"corrupt-{DateTime.UtcNow:yyyyMMddHHmmss}.config");
            File.Move(configPath, corruptPath, overwrite: true);
        }
    }

    // 設定ファイルを読み、XMLとして有効なら内容とハッシュを返す。
    // 設定の保存中でも読めるように、共有モードで開く。
    private static bool TryReadConfig(string path, out byte[] bytes, out string hash)
    {
        bytes = null;
        hash = null;
        try
        {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var ms = new MemoryStream();
            fs.CopyTo(ms);
            bytes = ms.ToArray();

            XDocument doc = XDocument.Load(new MemoryStream(bytes));
            if (doc.Root == null || doc.Root.Name.LocalName != "configuration")
            {
                bytes = null;
                return false;
            }
            hash = Convert.ToHexString(SHA256.HashData(bytes));
            return true;
        }
        catch
        {
            bytes = null;
            hash = null;
            return false;
        }
    }

    // 一時ファイルに書いてから置き換え、書き込み途中で止まっても不完全なファイルを残さない
    private static void WriteAtomic(string path, byte[] bytes)
    {
        string tempPath = path + ".tmp";
        File.WriteAllBytes(tempPath, bytes);
        File.Move(tempPath, path, overwrite: true);
    }

    private static BackupState LoadState()
    {
        try
        {
            string statePath = Path.Combine(backupDir, StateFileName);
            if (File.Exists(statePath))
            {
                return JsonSerializer.Deserialize<BackupState>(File.ReadAllText(statePath, Encoding.UTF8)) ?? new BackupState();
            }
        }
        catch (Exception ex)
        {
            Trace.WriteLine("バックアップの状態を読み込めませんでした: " + ex.Message);
        }
        return new BackupState();
    }

    private static void SaveState()
    {
        try
        {
            WriteAtomic(Path.Combine(backupDir, StateFileName),
                Encoding.UTF8.GetBytes(JsonSerializer.Serialize(state, JsonOptions)));
        }
        catch (Exception ex)
        {
            Trace.WriteLine("バックアップの状態を保存できませんでした: " + ex.Message);
        }
    }
}
