using System.Globalization;
using System.Text.RegularExpressions;

// 設定項目の「検証ルール」（キー名・型・範囲・許容値）。
// 既定値はここに持たない。既定値はひな形ファイル（AlwaysIME.dll.config.default）だけが持つ。
// ひな形と、利用者の設定ファイルの両方を、同じルールで検証する。
static class ConfigSchema
{
    private enum Kind
    {
        Text,        // 自由な文字列（空欄可）
        Int,         // 整数 Min～Max
        Float,       // 小数 Min～Max
        OneOf,       // Allowed のいずれか（大文字小文字は区別しない）
        IntPair,     // "a,b"（各値は 0～Max）
        RegexList,   // カンマ区切りの正規表現（各項目が正規表現として正しいこと。空欄可）
    }

    private sealed class Rule
    {
        public Kind Kind;
        public double Min;
        public double Max;
        public bool AllowEmpty;
        public string[] Allowed;
    }

    private static Rule Text() => new Rule { Kind = Kind.Text, AllowEmpty = true };
    private static Rule Int(int min, int max, bool allowEmpty = false) => new Rule { Kind = Kind.Int, Min = min, Max = max, AllowEmpty = allowEmpty };
    private static Rule Float(double min, double max, bool allowEmpty = false) => new Rule { Kind = Kind.Float, Min = min, Max = max, AllowEmpty = allowEmpty };
    private static Rule OneOf(bool allowEmpty, params string[] allowed) => new Rule { Kind = Kind.OneOf, Allowed = allowed, AllowEmpty = allowEmpty };
    private static Rule RegexList() => new Rule { Kind = Kind.RegexList, AllowEmpty = true };
    private static Rule IntPair(int max) => new Rule { Kind = Kind.IntPair, Max = max, AllowEmpty = true };

    // 設定ファイルの全キー（表示・補完・書き戻しの順序もこの並び）
    private static readonly (string Key, Rule Rule)[] Rules =
    {
        ("AlwaysIMEMode", OneOf(false, "1", "2", "3", "on", "off")),   // on/off は旧バージョンとの互換
        ("IsDarkMode", OneOf(false, "on", "off")),
        ("intervalTime", Int(10, 6000)),                                // ミリ秒。Timer.Intervalは0より大きい必要がある
        ("AppExitEnd", OneOf(false, "0", "1")),
        ("PassList", Text()),
        ("ImeOffList", Text()),
        ("ImeOffTitle", RegexList()),
        ("UpdateTag", RegexList()),
        ("NoKeyInputTime", Float(0, 1440)),                             // 分。(int)(分*60*1000)が桁あふれしない範囲
        ("SuspendFewTime", Int(0, 1440)),                               // 分
        ("SuspendTime", Int(0, 1440)),                                  // 分
        ("MsImeSwitch", OneOf(false, "0", "1")),
        ("Punctuation", IntPair(3)),                                    // 空欄 = 切り替えない
        ("SpaceWidth", IntPair(2)),                                     // 空欄 = 切り替えない
        ("KeyAssignment", OneOf(true, "0", "1")),                       // 空欄 = 管理しない
        ("BackupStableHours", Float(0, 8760)),
        ("BackupKeep", Int(1, 50)),
        ("LogEnabled", OneOf(false, "0", "1", "on", "off", "true", "false")),
        ("LogLevel", OneOf(false, "info", "warn", "error")),
        ("LogMaxSizeMB", Int(AppLog.MinMaxSizeMB, AppLog.MaxMaxSizeMB)),
        ("LogKeep", Int(AppLog.MinKeep, AppLog.MaxKeep)),
    };

    private static readonly Dictionary<string, Rule> RuleMap = BuildMap();

    private static Dictionary<string, Rule> BuildMap()
    {
        var map = new Dictionary<string, Rule>(StringComparer.Ordinal);
        foreach (var (key, rule) in Rules)
        {
            map.Add(key, rule);
        }
        return map;
    }

    // 全キー名（定義順）
    public static IEnumerable<string> Keys
    {
        get
        {
            foreach (var (key, _) in Rules)
            {
                yield return key;
            }
        }
    }

    public static bool IsKnownKey(string key) => key != null && RuleMap.ContainsKey(key);

    // 値が正しければ null、不正なら理由（ログ用。キー名と値は呼び出し側で付ける）を返す
    public static string Check(string key, string value)
    {
        if (!RuleMap.TryGetValue(key, out Rule rule))
        {
            return "未知のキーです";
        }
        value ??= "";
        if (value.Length == 0)
        {
            return rule.AllowEmpty ? null : "空欄にできません";
        }

        switch (rule.Kind)
        {
            case Kind.Text:
                return null;

            case Kind.Int:
                if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int i))
                {
                    return "整数ではありません";
                }
                return i >= rule.Min && i <= rule.Max ? null : $"範囲外です（{rule.Min}～{rule.Max}）";

            case Kind.Float:
                if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double d)
                    || double.IsNaN(d) || double.IsInfinity(d))
                {
                    return "数値ではありません";
                }
                return d >= rule.Min && d <= rule.Max ? null : $"範囲外です（{rule.Min}～{rule.Max}）";

            case Kind.OneOf:
                foreach (string allowed in rule.Allowed)
                {
                    if (string.Equals(value, allowed, StringComparison.OrdinalIgnoreCase))
                    {
                        return null;
                    }
                }
                return "許容値は " + string.Join(" / ", rule.Allowed) + " です";

            case Kind.IntPair:
                string[] parts = value.Split(',');
                if (parts.Length != 2)
                {
                    return "「数値,数値」の形式ではありません";
                }
                foreach (string part in parts)
                {
                    if (!int.TryParse(part, NumberStyles.Integer, CultureInfo.InvariantCulture, out int n) || n < 0 || n > rule.Max)
                    {
                        return $"各値は0～{rule.Max}の整数にしてください";
                    }
                }
                return null;

            case Kind.RegexList:
                // アプリと同じ分割（カンマ区切り、空の項目は無視）で、各項目を正規表現として検証する
                foreach (string pattern in value.Split(',', StringSplitOptions.RemoveEmptyEntries))
                {
                    try
                    {
                        _ = new Regex(pattern);
                    }
                    catch (ArgumentException ex)
                    {
                        return $"正規表現として不正です: {pattern}（{ex.Message}）";
                    }
                }
                return null;

            default:
                return null;
        }
    }
}
