using System.Diagnostics;
using System.Runtime.InteropServices;

// 「カタカナ・ひらがな・ローマ字」キーが押されたことを、どのアプリの操作中でも検出する。
//  ・キーの入力は妨げない（IME本来の動作はそのまま行われる）。見ているだけ
//  ・低レベルキーボードフック（WH_KEYBOARD_LL）を、専用スレッドで受ける。
//    UIスレッドで受けると、UIが止まっている間（設定ファイルをエディタで開いて待つ処理など）に
//    OSの応答待ち（既定300ms）を超えて、フックが外されてしまうため
//  ・押しっぱなしの自動リピートは、押した回数に数えない
//  ・管理者権限で動いているウィンドウの操作中は、OSの仕様でキーが届かない
//    （AlwaysIMEも管理者権限で起動していれば検出できる）
internal sealed class KanaKeyWatcher : IDisposable
{
    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN = 0x0100;
    private const int WM_KEYUP = 0x0101;
    private const int WM_SYSKEYDOWN = 0x0104;   // Altを押しながらの場合
    private const int WM_SYSKEYUP = 0x0105;
    private const uint WM_QUIT = 0x0012;
    private const uint PM_NOREMOVE = 0x0000;

    private const uint VK_KANA = 0x15;          // カタカナひらがなキー
    private const uint VK_OEM_COPY = 0xF2;      // 同じキーが、環境によってこの仮想キーで届くことがある
    private const uint SC_KANA = 0x70;          // 日本語キーボードのスキャンコード（仮想キーに依らない判定）
    private const uint LLKHF_EXTENDED = 0x01;

    // 自動リピートの間隔は、最初の1回だけ最大1秒ほど空く。これより長く何も届かなければ、
    // キーを離したのにその通知を受け取れなかった（UAC画面への切り替え中など）とみなして、押し直しとして扱う
    private const long RepeatGapMs = 1500;
    // 1回の押下が、続けて2回届いた場合の保険
    private const long DebounceMs = 150;

    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct MSG
    {
        public IntPtr hwnd;
        public uint message;
        public UIntPtr wParam;
        public IntPtr lParam;
        public uint time;
        public int ptX;
        public int ptY;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll")]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern int GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PeekMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax, uint wRemoveMsg);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool PostThreadMessage(uint idThread, uint msg, UIntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr GetModuleHandle(string lpModuleName);

    private readonly Action onPressed;
    private readonly LowLevelKeyboardProc proc;     // フックの間、GCに回収されないよう保持しておく
    private readonly ManualResetEventSlim ready = new ManualResetEventSlim(false);
    private Thread thread;
    private uint threadId;
    private IntPtr hook;

    // フックのスレッドだけが触る状態
    private bool kanaIsDown;
    private long lastDownTick;
    private long lastTriggerTick;

    // onPressed は、フックのスレッドから呼ばれる。すぐ戻ること（UIの処理は BeginInvoke などで渡す）
    public KanaKeyWatcher(Action onPressed)
    {
        this.onPressed = onPressed ?? throw new ArgumentNullException(nameof(onPressed));
        proc = HookProc;
    }

    public void Start()
    {
        if (thread != null)
        {
            return;
        }
        thread = new Thread(Run) { IsBackground = true, Name = "KanaKeyWatcher" };
        thread.Start();
    }

    private void Run()
    {
        threadId = GetCurrentThreadId();
        // メッセージキューを先に作っておく（Dispose が WM_QUIT を送れるようにする）
        PeekMessage(out _, IntPtr.Zero, 0, 0, PM_NOREMOVE);
        try
        {
            hook = SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(null), 0);
            if (hook == IntPtr.Zero)
            {
                AppLog.Warn($"カタカナ・ひらがなキーの監視を開始できませんでした（エラー {Marshal.GetLastWin32Error()}）");
                return;
            }
            AppLog.Info("カタカナ・ひらがなキーの監視を開始しました");
        }
        finally
        {
            ready.Set();
        }

        try
        {
            // フックの呼び出しは、この待機の中で処理される。WM_QUIT で抜ける
            while (GetMessage(out _, IntPtr.Zero, 0, 0) > 0)
            {
            }
        }
        finally
        {
            UnhookWindowsHookEx(hook);
            hook = IntPtr.Zero;
        }
    }

    private IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0)
        {
            try
            {
                // KBDLLHOOKSTRUCT: vkCode(0) scanCode(4) flags(8)
                uint vk = (uint)Marshal.ReadInt32(lParam, 0);
                uint scan = (uint)Marshal.ReadInt32(lParam, 4);
                uint flags = (uint)Marshal.ReadInt32(lParam, 8);
                if (IsKanaKey(vk, scan, flags))
                {
                    int message = wParam.ToInt32();
                    if (message == WM_KEYDOWN || message == WM_SYSKEYDOWN)
                    {
                        OnKanaDown();
                    }
                    else if (message == WM_KEYUP || message == WM_SYSKEYUP)
                    {
                        kanaIsDown = false;
                    }
                }
            }
            catch
            {
                // フックの中では、例外を外へ出さない（キー入力を妨げないため）
            }
        }
        // 常に次へ渡す。キーは消費しない
        return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam);
    }

    private static bool IsKanaKey(uint vk, uint scan, uint flags)
    {
        return vk == VK_KANA
            || vk == VK_OEM_COPY
            || (scan == SC_KANA && (flags & LLKHF_EXTENDED) == 0);
    }

    private void OnKanaDown()
    {
        long now = Environment.TickCount64;
        bool repeat = kanaIsDown && now - lastDownTick < RepeatGapMs;
        kanaIsDown = true;
        lastDownTick = now;
        if (repeat)
        {
            return; // 押しっぱなしの自動リピート
        }
        if (lastTriggerTick != 0 && now - lastTriggerTick < DebounceMs)
        {
            return;
        }
        lastTriggerTick = now;
        onPressed();
    }

    public void Dispose()
    {
        Thread t = thread;
        if (t == null)
        {
            return;
        }
        thread = null;
        if (ready.Wait(1000))
        {
            PostThreadMessage(threadId, WM_QUIT, UIntPtr.Zero, IntPtr.Zero);
        }
        t.Join(1000);
    }
}
