using Microsoft.Win32;
using System.Text;

// 互換モード（旧IME）のMicrosoft IMEで、変換キー→IME-ON / 「無変換キー」→IME-OFF を割り当てる。
//
// 新しいMS-IME（標準モード）は MSIME キーのDWORD 3つ（IsKeyAssignmentEnabled など）で割り当てるが、
// 旧IMEは「キー設定テーブル」を持っていて、割り当てはそのテーブルの書き換えになる。
//
// 実機のレジストリ出力4点（初期値 / 割り当て / 解除 / 再割り当て）の差分から、
// 旧IMEのプロパティ画面が行う変更は次の3つだけだった（HKCU\Software\Microsoft\IME\15.0\IMEJP 配下）。
//   1. MSIME\keystyle を "Custom" にする（初期値は "NATURAL"）
//   2. StyleList\Custom\key   の「変換」の先頭の機能コードを CE（IME-ON）、「無変換」を CF（IME-OFF）にする
//   3. StyleList\Custom\S4key に「変換=CE 00 00 00 00 00」を置く
// 解除すると、変換は 87（再変換）、無変換は 97（かな切替）に戻す（key と、S4key の変換）。
// 87 / 97 は、プリセット（NATURAL など）の初期値と同じ（旧IME環境のレジストリ出力で確認）。
// ※旧IMEの画面の「解除」は 80（割り当てなし）にして S4key の項目を消すが、ここでは上の値に戻す。
// 初回だけは、プリセット（NATURAL）から StyleList\Custom が丸ごと作られる。
//
// テーブルの形式: REG_BINARY。「キー名=XX XX XX XX XX XX」をNULで区切り、最後にNULをもう1つ付ける。
// キー名はShift-JIS。「XX」は6つの入力状態ごとの機能コード（先頭＝入力文字なしの状態）。
//
// 設計方針
//  ・終了時に元へ戻す（新IME側の RestoreKeyAssignment と同じ考え方）。
//    ただし、実行中に利用者がIMEの設定画面で変更していたら、その変更を壊さないよう戻さない。
//  ・想定外の形（別名形式の項目、Custom が残っているが有効でない、など）は何も変更せず、ログだけ残す。
static class LegacyKeyAssignment
{
    private const string MsImeKeyPath = @"Software\Microsoft\IME\15.0\IMEJP\MSIME";
    private const string StyleListPath = @"Software\Microsoft\IME\15.0\IMEJP\StyleList";
    private const string CustomStyle = "Custom";
    private const string DefaultPreset = "NATURAL"; // keystyle が無いときの既定

    // Shift-JIS のキー名
    private static readonly byte[] Henkan = { 0x95, 0xCF, 0x8A, 0xB7 };                // 変換
    private static readonly byte[] Muhenkan = { 0x96, 0xB3, 0x95, 0xCF, 0x8A, 0xB7 };  // 無変換

    private const string CodeImeOn = "CE";   // IME-オン
    private const string CodeImeOff = "CF";  // IME-オフ
    private const string CodeReconvert = "87"; // 再変換（変換キーの初期値）
    private const string CodeKanaSwitch = "97"; // かな切替（「無変換キー」の初期値）
    private const string CodeNone = "80";       // 割り当てなし（画面の「解除」で入る値）

    // 画面での変更のたびに MSIME\SerialNo が「起動からのミリ秒」で更新されていた（IME側が変更検知に使っている可能性）。
    // 実機で、書き換えが即時に反映されないときに true にして確認する。反映されるなら false のままでよい。
    private const bool UpdateSerialNo = true;

    // 実行中に行った変更（終了時の復元用）
    private sealed class Session
    {
        public string OrigKeyStyle;     // null = 値が無かった
        public bool CustomExisted;
        public byte[] OrigKey;          // null = 値が無かった
        public byte[] OrigS4Key;
        // 最後に自分が書いた状態（復元の前に、利用者が変更していないかを確かめるため）
        public string WrittenKeyStyle;
        public byte[] WrittenKey;
        public byte[] WrittenS4Key;
    }

    private static Session session;

    // ---- 公開 ----

    // enable=true: 割り当てる / false: 割り当てを外す。成功（または何もする必要が無い）なら true
    public static bool Apply(bool enable)
    {
        try
        {
            return ApplyCore(enable);
        }
        catch (Exception ex)
        {
            AppLog.Warn("旧IMEのキー割り当てに失敗しました", ex);
            Restore(force: true); // 途中まで書いた分を戻す
            return false;
        }
    }

    // 変更前の状態に戻す。利用者が実行中に設定を変えていた場合は戻さない
    public static void Restore() => Restore(force: false);

    // 現在、変換→IME-ON / 無変換→IME-OFF になっているか（設定画面の初期表示用）
    public static bool IsAssigned()
    {
        try
        {
            using (RegistryKey msime = Registry.CurrentUser.OpenSubKey(MsImeKeyPath))
            using (RegistryKey custom = Registry.CurrentUser.OpenSubKey(StyleListPath + @"\" + CustomStyle))
            {
                if (!string.Equals(msime?.GetValue("keystyle") as string, CustomStyle, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
                List<byte[]> key = ParseTable(custom?.GetValue("key") as byte[]);
                return IsCode(key, Henkan, CodeImeOn) && IsCode(key, Muhenkan, CodeImeOff);
            }
        }
        catch (Exception ex)
        {
            AppLog.Warn("旧IMEのキー割り当ての確認に失敗しました", ex);
            return false;
        }
    }

    // ---- 割り当て ----

    private static bool ApplyCore(bool enable)
    {
        using (RegistryKey msime = Registry.CurrentUser.OpenSubKey(MsImeKeyPath, writable: true))
        using (RegistryKey styleList = Registry.CurrentUser.OpenSubKey(StyleListPath, writable: true))
        {
            if (msime == null || styleList == null)
            {
                AppLog.Warn("旧IMEのキー設定（StyleList）が見つかりません");
                return false;
            }

            string activeStyle = msime.GetValue("keystyle") as string;
            if (string.IsNullOrEmpty(activeStyle))
            {
                activeStyle = DefaultPreset;
            }
            bool activeIsCustom = string.Equals(activeStyle, CustomStyle, StringComparison.OrdinalIgnoreCase);
            bool customExists;
            using (RegistryKey probe = styleList.OpenSubKey(CustomStyle))
            {
                customExists = probe != null;
            }

            if (!enable)
            {
                // 外す対象（Custom の中の IME-ON/OFF）があるときだけ処理する
                if (!activeIsCustom || !customExists)
                {
                    return true;
                }
                BeginSession(msime, styleList);
                return EditCustom(msime, styleList, assign: false);
            }

            if (!activeIsCustom)
            {
                if (customExists)
                {
                    // 現在のキー設定（NATURAL など）と、残っている Custom の内容は別物。
                    // 勝手に Custom へ切り替えると他のキー動作まで変わるので、何もしない
                    AppLog.Warn($"Custom が残っていますが、現在のキー設定は {activeStyle} のため、割り当てを行いません");
                    return false;
                }
                BeginSession(msime, styleList);
                if (!CreateCustomFromPreset(styleList, activeStyle))
                {
                    session = null;
                    return false;
                }
            }
            else
            {
                BeginSession(msime, styleList);
            }

            if (!EditCustom(msime, styleList, assign: true))
            {
                if (!activeIsCustom)
                {
                    Restore(force: true); // 作りかけの Custom を片付ける
                }
                return false;
            }

            if (!activeIsCustom)
            {
                // テーブルを書き終えてから切り替える（IMEが中途半端なテーブルを見ないように、最後に行う）
                msime.SetValue("keystyle", CustomStyle, RegistryValueKind.String);
                session.WrittenKeyStyle = CustomStyle;
                BumpSerialNo(msime);
            }
            AppLog.Info("旧IMEのキー割り当て（変換→IME-ON / 無変換→IME-OFF）を有効にしました");
            return true;
        }
    }

    // Custom の key / S4key を書き換える。変更があれば書き込み、session に「最後に書いた状態」を記録する
    private static bool EditCustom(RegistryKey msime, RegistryKey styleList, bool assign)
    {
        using (RegistryKey custom = styleList.OpenSubKey(CustomStyle, writable: true))
        {
            byte[] keyBytes = custom?.GetValue("key") as byte[];
            if (keyBytes == null)
            {
                AppLog.Warn("Custom に key がありません");
                return false;
            }
            byte[] s4Bytes = custom.GetValue("S4key") as byte[]; // 無ければ空のテーブルとして扱う

            List<byte[]> key = ParseTable(keyBytes);
            List<byte[]> s4 = ParseTable(s4Bytes);
            bool ok;
            bool keyChanged1 = false, keyChanged2 = false, s4Changed = false;
            if (assign)
            {
                ok = SetFirstCode(key, Henkan, CodeImeOn, out keyChanged1)
                  && SetFirstCode(key, Muhenkan, CodeImeOff, out keyChanged2)
                  && SetFirstCode(s4, Henkan, CodeImeOn, out s4Changed);
            }
            else
            {
                // 変換→再変換 / 無変換→かな切替 に戻す。IME-ON/OFF（または割り当てなし）のときだけ書き換え、
                // 他の機能を割り当てている場合は触らない。S4key の変換は、項目が無ければ作る
                ok = ReplaceFirstCode(key, Henkan, CodeImeOn, CodeReconvert, addIfMissing: false, out keyChanged1)
                  && ReplaceFirstCode(key, Muhenkan, CodeImeOff, CodeKanaSwitch, addIfMissing: false, out keyChanged2)
                  && ReplaceFirstCode(s4, Henkan, CodeImeOn, CodeReconvert, addIfMissing: true, out s4Changed);
            }
            if (!ok)
            {
                AppLog.Warn("変換/無変換が別名形式（1項目に複数のキー名）で定義されているため、変更しません");
                return false;
            }

            byte[] newKey = keyBytes;
            byte[] newS4 = s4Bytes;
            if (keyChanged1 || keyChanged2)
            {
                newKey = BuildTable(key);
                custom.SetValue("key", newKey, RegistryValueKind.Binary);
            }
            if (s4Changed)
            {
                newS4 = BuildTable(s4);
                custom.SetValue("S4key", newS4, RegistryValueKind.Binary);
            }
            session.WrittenKey = newKey;
            session.WrittenS4Key = newS4;
            if (keyChanged1 || keyChanged2 || s4Changed)
            {
                BumpSerialNo(msime);
            }
            return true;
        }
    }

    // プリセット（NATURAL など）の値を Custom へ写す。画面が作る Custom と同じく、表示名と既定ローマ字表は写さない。
    // ※画面が作る Custom は、ここで作るものより項目が多い（画面側が内蔵の既定値を展開して書いている）。
    //   Custom の中身が不足して他のキーが効かなくなる場合は、実機のレジストリ出力から Custom 一式を埋め込む方式に変える。
    private static bool CreateCustomFromPreset(RegistryKey styleList, string presetName)
    {
        using (RegistryKey preset = styleList.OpenSubKey(presetName))
        {
            if (preset == null)
            {
                AppLog.Warn($"キー設定 {presetName} が見つからないため、Custom を作れません");
                return false;
            }
            using (RegistryKey custom = styleList.CreateSubKey(CustomStyle))
            {
                foreach (string name in preset.GetValueNames())
                {
                    if (name == "DisplayName" || name == "defRoma")
                    {
                        continue;
                    }
                    object value = preset.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                    custom.SetValue(name, value, preset.GetValueKind(name));
                }
            }
        }
        return true;
    }

    // ---- 復元 ----

    private static void BeginSession(RegistryKey msime, RegistryKey styleList)
    {
        if (session != null)
        {
            return; // 最初に記録した「変更前」を保つ
        }
        string keyStyle = msime.GetValue("keystyle") as string;
        using (RegistryKey custom = styleList.OpenSubKey(CustomStyle))
        {
            session = new Session
            {
                OrigKeyStyle = keyStyle,
                CustomExisted = custom != null,
                OrigKey = custom?.GetValue("key") as byte[],
                OrigS4Key = custom?.GetValue("S4key") as byte[],
                WrittenKeyStyle = keyStyle,
            };
            session.WrittenKey = session.OrigKey;
            session.WrittenS4Key = session.OrigS4Key;
        }
    }

    private static void Restore(bool force)
    {
        Session s = session;
        if (s == null)
        {
            return;
        }
        session = null;
        try
        {
            using (RegistryKey msime = Registry.CurrentUser.OpenSubKey(MsImeKeyPath, writable: true))
            using (RegistryKey styleList = Registry.CurrentUser.OpenSubKey(StyleListPath, writable: true))
            {
                if (msime == null || styleList == null)
                {
                    return;
                }

                byte[] curKey, curS4;
                using (RegistryKey custom = styleList.OpenSubKey(CustomStyle))
                {
                    curKey = custom?.GetValue("key") as byte[];
                    curS4 = custom?.GetValue("S4key") as byte[];
                }
                bool untouched = string.Equals(msime.GetValue("keystyle") as string, s.WrittenKeyStyle, StringComparison.OrdinalIgnoreCase)
                    && BytesEqual(curKey, s.WrittenKey)
                    && BytesEqual(curS4, s.WrittenS4Key);
                if (!force && !untouched)
                {
                    AppLog.Info("旧IMEのキー設定が実行中に変更されたため、元に戻しません");
                    return;
                }

                if (s.OrigKeyStyle == null)
                {
                    msime.DeleteValue("keystyle", throwOnMissingValue: false);
                }
                else
                {
                    msime.SetValue("keystyle", s.OrigKeyStyle, RegistryValueKind.String);
                }

                if (!s.CustomExisted)
                {
                    styleList.DeleteSubKeyTree(CustomStyle, throwOnMissingSubKey: false);
                }
                else
                {
                    using (RegistryKey custom = styleList.OpenSubKey(CustomStyle, writable: true))
                    {
                        if (custom != null)
                        {
                            RestoreBinary(custom, "key", s.OrigKey);
                            RestoreBinary(custom, "S4key", s.OrigS4Key);
                        }
                    }
                }
                BumpSerialNo(msime);
            }
        }
        catch (Exception ex)
        {
            AppLog.Warn("旧IMEのキー割り当てを元に戻せませんでした", ex);
        }
    }

    private static void RestoreBinary(RegistryKey key, string name, byte[] original)
    {
        if (original == null)
        {
            key.DeleteValue(name, throwOnMissingValue: false);
        }
        else
        {
            key.SetValue(name, original, RegistryValueKind.Binary);
        }
    }

    private static void BumpSerialNo(RegistryKey msime)
    {
        if (UpdateSerialNo)
        {
            msime.SetValue("SerialNo", unchecked((int)Environment.TickCount64), RegistryValueKind.DWord);
        }
    }

    private static bool BytesEqual(byte[] a, byte[] b)
    {
        if (a == null || b == null)
        {
            return a == null && b == null;
        }
        return a.AsSpan().SequenceEqual(b.AsSpan());
    }

    // ---- キー設定テーブル（NUL区切りの「キー名=XX XX XX XX XX XX」） ----

    private static List<byte[]> ParseTable(byte[] data)
    {
        var list = new List<byte[]>();
        if (data == null)
        {
            return list;
        }
        int pos = 0;
        while (pos < data.Length)
        {
            int end = Array.IndexOf(data, (byte)0, pos);
            if (end < 0)
            {
                list.Add(data[pos..]); // 終端のNULが無い末尾（通常は無い）
                break;
            }
            if (end == pos)
            {
                break; // 空の項目 = テーブルの終端
            }
            list.Add(data[pos..end]);
            pos = end + 1;
        }
        return list;
    }

    private static byte[] BuildTable(List<byte[]> entries)
    {
        using (var ms = new MemoryStream())
        {
            foreach (byte[] entry in entries)
            {
                ms.Write(entry, 0, entry.Length);
                ms.WriteByte(0);
            }
            ms.WriteByte(0);
            return ms.ToArray();
        }
    }

    // keyName の項目の位置を返す（無ければ -1）。
    // 「Ctrl+A ←」のように1項目に複数のキー名が空白区切りで並ぶ形式（別名）の中に keyName があれば alias = true
    private static int FindEntry(List<byte[]> table, byte[] keyName, out bool alias)
    {
        int found = -1;
        alias = false;
        for (int i = 0; i < table.Count; i++)
        {
            byte[] entry = table[i];
            int eq = Array.IndexOf(entry, (byte)'=');
            if (eq < 0)
            {
                continue;
            }
            ReadOnlySpan<byte> name = entry.AsSpan(0, eq);
            if (name.SequenceEqual(keyName.AsSpan()))
            {
                if (found < 0)
                {
                    found = i;
                }
            }
            else if (name.IndexOf((byte)' ') >= 0 && ContainsToken(name, keyName))
            {
                alias = true;
            }
        }
        return found;
    }

    private static bool ContainsToken(ReadOnlySpan<byte> name, byte[] token)
    {
        int start = 0;
        while (start <= name.Length)
        {
            ReadOnlySpan<byte> rest = name.Slice(start);
            int sp = rest.IndexOf((byte)' ');
            ReadOnlySpan<byte> part = sp < 0 ? rest : rest.Slice(0, sp);
            if (part.SequenceEqual(token.AsSpan()))
            {
                return true;
            }
            if (sp < 0)
            {
                break;
            }
            start += sp + 1;
        }
        return false;
    }

    // 項目の先頭の機能コード（「XX XX …」の最初の XX）
    private static string GetFirstCode(byte[] entry)
    {
        int start = Array.IndexOf(entry, (byte)'=') + 1;
        int end = Array.IndexOf(entry, (byte)' ', start);
        if (end < 0)
        {
            end = entry.Length;
        }
        return Encoding.ASCII.GetString(entry, start, end - start);
    }

    private static byte[] WithFirstCode(byte[] entry, string code)
    {
        int valueStart = Array.IndexOf(entry, (byte)'=') + 1;
        int restStart = Array.IndexOf(entry, (byte)' ', valueStart);
        if (restStart < 0)
        {
            restStart = entry.Length;
        }
        byte[] codeBytes = Encoding.ASCII.GetBytes(code);
        var result = new byte[valueStart + codeBytes.Length + (entry.Length - restStart)];
        Buffer.BlockCopy(entry, 0, result, 0, valueStart);
        Buffer.BlockCopy(codeBytes, 0, result, valueStart, codeBytes.Length);
        Buffer.BlockCopy(entry, restStart, result, valueStart + codeBytes.Length, entry.Length - restStart);
        return result;
    }

    private static byte[] NewEntry(byte[] keyName, string code)
    {
        byte[] tail = Encoding.ASCII.GetBytes("=" + code + " 00 00 00 00 00"); // 画面が書く形式に合わせる
        var result = new byte[keyName.Length + tail.Length];
        Buffer.BlockCopy(keyName, 0, result, 0, keyName.Length);
        Buffer.BlockCopy(tail, 0, result, keyName.Length, tail.Length);
        return result;
    }

    private static bool IsCode(List<byte[]> table, byte[] keyName, string code)
    {
        int idx = FindEntry(table, keyName, out _);
        return idx >= 0 && string.Equals(GetFirstCode(table[idx]), code, StringComparison.OrdinalIgnoreCase);
    }

    // keyName の先頭の機能コードを code にする（項目が無ければ追加）。別名形式で定義されている場合は false（変更しない）
    private static bool SetFirstCode(List<byte[]> table, byte[] keyName, string code, out bool changed)
    {
        changed = false;
        int idx = FindEntry(table, keyName, out bool alias);
        if (alias)
        {
            return false;
        }
        if (idx < 0)
        {
            table.Add(NewEntry(keyName, code));
            changed = true;
        }
        else if (!string.Equals(GetFirstCode(table[idx]), code, StringComparison.OrdinalIgnoreCase))
        {
            table[idx] = WithFirstCode(table[idx], code);
            changed = true;
        }
        return true;
    }

    // 先頭の機能コードが expected（または割り当てなし）のときだけ replacement にする。
    // 項目が無いとき: addIfMissing なら replacement で作り、そうでなければ何もしない。別名形式の場合は false（変更しない）
    private static bool ReplaceFirstCode(List<byte[]> table, byte[] keyName, string expected, string replacement, bool addIfMissing, out bool changed)
    {
        changed = false;
        int idx = FindEntry(table, keyName, out bool alias);
        if (alias)
        {
            return false;
        }
        if (idx < 0)
        {
            if (addIfMissing)
            {
                table.Add(NewEntry(keyName, replacement));
                changed = true;
            }
            return true;
        }
        string current = GetFirstCode(table[idx]);
        if (string.Equals(current, expected, StringComparison.OrdinalIgnoreCase)
            || string.Equals(current, CodeNone, StringComparison.OrdinalIgnoreCase))
        {
            table[idx] = WithFirstCode(table[idx], replacement);
            changed = true;
        }
        return true;
    }
}
