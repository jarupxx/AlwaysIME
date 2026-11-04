using System.Runtime.InteropServices;
using Microsoft.Win32;

// 現在選択中の入力方式（IME）が、「MS-IME入力設定」に対応しているかを調べる。
//  ・Windows 11の新しいMS-IME（標準モード）だけが対応する
//  ・互換モードのMS-IME（旧IME）や、ATOK、Google日本語入力などは対応しない
// 判定できないとき（取得に失敗した、IMEではなくキーボードレイアウトが選ばれている など）は Unknown を返し、何も表示しない。
static class ImeEnvironment
{
    public enum Status
    {
        Unknown,    // 判定できない（警告は出さない）
        Supported,  // 新しいMS-IME（標準モード）
        CompatMode, // 互換モードのMS-IME（旧IME）
        NotMsIme,   // MS-IME以外のIME
    }

    // MS-IMEのTIP（テキストサービス）のCLSID。互換モードの設定キー名にも使われる
    private static readonly Guid MsImeClsid = new Guid("03b5835f-f03c-411b-9ce2-aa23e1171e36");
    private static readonly Guid CategoryKeyboard = new Guid("34745c63-b2f0-4784-8b67-5e12c8701a31"); // GUID_TFCAT_TIP_KEYBOARD
    private static readonly Guid ClsidInputProcessorProfiles = new Guid("33c53a50-f456-4884-b049-85fd643ecfed");
    private const uint TF_PROFILETYPE_INPUTPROCESSOR = 1;
    // 新しいMS-IME（標準モード）が導入されたのは Windows 10 2004（ビルド19041）から。それより前のOS（Server 2019 = 17763 など）は旧IMEのみ
    private const int FirstBuildWithNewMsIme = 19041;
    private const string ImeJpKeyPath = @"Software\Microsoft\IME\15.0\IMEJP"; // 日本語IMEのユーザー設定

    [StructLayout(LayoutKind.Sequential)]
    private struct TF_INPUTPROCESSORPROFILE
    {
        public uint dwProfileType;
        public ushort langid;
        public Guid clsid;
        public Guid catid;
        public IntPtr hklSubstitute;
        public uint dwCaps;
        public Guid guidProfile;
        public IntPtr hkl;
        public uint dwFlags;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct RTL_OSVERSIONINFOW
    {
        public uint dwOSVersionInfoSize;
        public uint dwMajorVersion;
        public uint dwMinorVersion;
        public uint dwBuildNumber;
        public uint dwPlatformId;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string szCSDVersion;
    }

    // マニフェストの有無に左右されず、実際のビルド番号を返す（Environment.OSVersion は古い値を返すことがある）
    [DllImport("ntdll.dll")]
    private static extern int RtlGetVersion(ref RTL_OSVERSIONINFOW versionInfo);

    // GetActiveProfile までの並び（vtableの順）を合わせるため、前のメソッドも宣言している
    [ComImport, Guid("71c6e74c-0f28-11d8-a82a-00065b84435c"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface ITfInputProcessorProfileMgr
    {
        void ActivateProfile(uint profileType, ushort langid, ref Guid clsid, ref Guid guidProfile, IntPtr hkl, uint flags);
        void DeactivateProfile(uint profileType, ushort langid, ref Guid clsid, ref Guid guidProfile, IntPtr hkl, uint flags);
        void GetProfile(uint profileType, ushort langid, ref Guid clsid, ref Guid guidProfile, IntPtr hkl, out TF_INPUTPROCESSORPROFILE profile);
        void EnumProfiles(ushort langid, out IntPtr enumerator);
        void ReleaseInputProcessor(ref Guid clsid, uint flags);
        void RegisterProfile(ref Guid clsid, ushort langid, ref Guid guidProfile,
            [MarshalAs(UnmanagedType.LPWStr)] string description, uint descriptionLength,
            [MarshalAs(UnmanagedType.LPWStr)] string iconFile, uint iconFileLength, uint iconIndex,
            IntPtr hklSubstitute, uint caps, int enabledByDefault, uint flags);
        void UnregisterProfile(ref Guid clsid, ushort langid, ref Guid guidProfile, uint flags);
        void GetActiveProfile(ref Guid catid, out TF_INPUTPROCESSORPROFILE profile);
    }

    // UIスレッド（STA）から呼ぶこと。例外は外へ出さない
    public static Status GetStatus()
    {
        try
        {
            Guid active;
            if (!TryGetActiveTipClsid(out active))
            {
                return Status.Unknown;
            }
            if (active != MsImeClsid)
            {
                return Status.NotMsIme;
            }
            return IsCompatMode() ? Status.CompatMode : Status.Supported;
        }
        catch
        {
            return Status.Unknown;
        }
    }

    // 現在選択中の入力方式が、IME（TIP）ならそのCLSIDを返す。キーボードレイアウトの場合は false
    private static bool TryGetActiveTipClsid(out Guid clsid)
    {
        clsid = Guid.Empty;
        object com = null;
        try
        {
            Type type = Type.GetTypeFromCLSID(ClsidInputProcessorProfiles);
            if (type == null)
            {
                return false;
            }
            com = Activator.CreateInstance(type);
            var mgr = (ITfInputProcessorProfileMgr)com;
            Guid category = CategoryKeyboard;
            mgr.GetActiveProfile(ref category, out TF_INPUTPROCESSORPROFILE profile);
            if (profile.dwProfileType != TF_PROFILETYPE_INPUTPROCESSOR)
            {
                return false;
            }
            clsid = profile.clsid;
            return clsid != Guid.Empty;
        }
        finally
        {
            if (com != null && Marshal.IsComObject(com))
            {
                Marshal.ReleaseComObject(com);
            }
        }
    }

    // 旧IMEか標準モード（新IME）かを判定する。互換モード（旧IME）なら true
    //  1. NoTsf3Override2 があれば、その値に従う（1: 互換モード、それ以外: 標準モード）
    //  2. 値が無い場合は、OSのビルド番号で判断する
    //     ・19041未満（Windows Server 2019など）は新IMEが存在しないので旧IME
    //     ・19041以上は、互換モードに切り替えていない限り新IME
    //  3. ビルド番号を取得できなかった場合だけ、レジストリのキー構成から判断する（LooksLikeLegacyImeProfile）
    private static bool IsCompatMode()
    {
        using (RegistryKey key = Registry.CurrentUser.OpenSubKey(
            @"Software\Microsoft\Input\TSF\Tsf3Override\{" + MsImeClsid.ToString() + "}"))
        {
            if (key?.GetValue("NoTsf3Override2") is int value)
            {
                return value == 1;
            }
        }

        int build = GetOsBuildNumber();
        if (build > 0)
        {
            return build < FirstBuildWithNewMsIme;
        }
        return LooksLikeLegacyImeProfile();
    }

    // OSのビルド番号。取得できなければ 0
    private static int GetOsBuildNumber()
    {
        try
        {
            var info = new RTL_OSVERSIONINFOW();
            info.dwOSVersionInfoSize = (uint)Marshal.SizeOf(typeof(RTL_OSVERSIONINFOW));
            return RtlGetVersion(ref info) == 0 ? (int)info.dwBuildNumber : 0;
        }
        catch
        {
            return 0;
        }
    }

    // 旧IMEのプロファイルか（ビルド番号が取得できないときの予備判定）。
    // 実機のレジストリ出力（HKCU\Software\Microsoft\IME\15.0\IMEJP）の比較から、
    //   旧IME（Server 2019）      : Colors あり、ResetExcludeProperties なし
    //   新IME（Win10 22h2 / Win11）: Colors なし、ResetExcludeProperties あり
    // であることを根拠にしている。どちらか一方だけでは判断せず、両方が旧IMEの特徴に一致したときだけ true にする
    // （新IMEを選んだばかりでキーが未作成の環境を、誤って旧IMEと判定しないため）
    private static bool LooksLikeLegacyImeProfile()
    {
        using (RegistryKey imeJp = Registry.CurrentUser.OpenSubKey(ImeJpKeyPath))
        {
            if (imeJp == null)
            {
                return false;
            }
            bool hasLegacyMarker = HasSubKey(imeJp, "Colors");
            bool hasNewMarker = HasSubKey(imeJp, "ResetExcludeProperties");
            return hasLegacyMarker && !hasNewMarker;
        }
    }

    private static bool HasSubKey(RegistryKey parent, string name)
    {
        using (RegistryKey sub = parent.OpenSubKey(name))
        {
            return sub != null;
        }
    }
}
