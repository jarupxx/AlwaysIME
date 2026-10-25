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

    // 互換モード（旧IME）か。NoTsf3Override2 が 1 なら互換モード。0、または値やキーが無ければ標準モード（Windows 10 / 11共通）
    private static bool IsCompatMode()
    {
        using RegistryKey key = Registry.CurrentUser.OpenSubKey(
            @"Software\Microsoft\Input\TSF\Tsf3Override\{" + MsImeClsid.ToString() + "}");
        return key?.GetValue("NoTsf3Override2") is int value && value == 1;
    }
}
