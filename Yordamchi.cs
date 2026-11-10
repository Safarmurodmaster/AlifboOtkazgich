using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace AlifboYordamchi
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            bool createdNew;
            using (Mutex mutex = new Mutex(true, "AlifboYordamchi_SingleInstance_Mutex", out createdNew))
            {
                if (!createdNew)
                {
                    return;
                }

                // TLS 1.2 ni yoqish (GitHub HTTPS yuklab olish uchun kerak)
                try
                {
                    ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072 | SecurityProtocolType.Tls;
                }
                catch { }

                // Kuchsiz kompyuterlarda klaviatura xabarlarini zudlik bilan qabul qilish
                try
                {
                    Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.AboveNormal;
                }
                catch { }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new TrayApplicationContext());
            }
        }
    }

    public class TrayApplicationContext : ApplicationContext
    {
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;
        private ToolStripMenuItem itemActive;
        private bool isActive = true;
        private bool isLicensedCached = false;

        // Versiya
        public const string CURRENT_VERSION = "1.0.4";
        // GitHub orqali versiya tekshirish URL manzili
        public const string GITHUB_VERSION_URL = "https://raw.githubusercontent.com/Safarmurodmaster/AlifboOtkazgich/main/version.json";
        public const string GITHUB_SETUP_URL = "https://raw.githubusercontent.com/Safarmurodmaster/AlifboOtkazgich/main/AlifboOtkazgich-Setup.exe";

        private const int WM_HOTKEY = 0x0312;
        private const int HOTKEY_PASTE = 1;
        private const int HOTKEY_TOGGLE = 2;

        private const uint MOD_ALT = 0x0001;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;

        private const string SALT = "SAFAR_2026_";
        private static readonly IntPtr MAGIC_EXTRA = (IntPtr)0x414C4642; // "ALFB" o'zimizning signallarimiz uchun

        #region Win32 API va Keyboard Hook

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int nVirtKey);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public uint type;
            public InputUnion U;
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct InputUnion
        {
            [FieldOffset(0)] public MOUSEINPUT mi;
            [FieldOffset(0)] public KEYBDINPUT ki;
            [FieldOffset(0)] public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const uint LLKHF_INJECTED = 0x10;

        private const uint INPUT_KEYBOARD = 1;
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_UNICODE = 0x0004;

        private const int VK_BACK = 0x08;
        private const int VK_TAB = 0x09;
        private const int VK_RETURN = 0x0D;
        private const int VK_SHIFT = 0x10;
        private const int VK_CONTROL = 0x11;
        private const int VK_MENU = 0x12; // Alt
        private const int VK_CAPITAL = 0x14; // CapsLock
        private const int VK_ESCAPE = 0x1B;
        private const int VK_SPACE = 0x20;
        private const int VK_OEM_3 = 0xC0; // ` yoki ~
        private const int VK_OEM_7 = 0xDE; // ' yoki "

        private IntPtr hookId = IntPtr.Zero;
        private LowLevelKeyboardProc hookProc;
        private char lastChar = '\0';
        private long lastCharTime = 0;

        // Tezkor bufer rejimi (Backspace ishlatmaslik va sş, oö xatolarini 100% yo'qotish uchun)
        private System.Windows.Forms.Timer bufferTimer;
        private char bufferedChar = '\0';
        private bool useBufferMode = true;

        #endregion

        private HiddenWindow hiddenWindow;
        private System.Windows.Forms.Timer updateCheckTimer;
        private System.Windows.Forms.Timer watchdogTimer;
        private DateTime lastKeyTime = DateTime.MinValue;

        [DllImport("kernel32.dll")]
        private static extern bool SetProcessWorkingSetSize(IntPtr proc, int min, int max);

        public TrayApplicationContext()
        {
            isLicensedCached = IsLicensed();

            // Harflarni ekranga chiqarishdan oldin qisqa ushlab turuvchi (200 ms) bufer taymeri
            bufferTimer = new System.Windows.Forms.Timer();
            bufferTimer.Interval = 200;
            bufferTimer.Tick += (s, e) =>
            {
                bufferTimer.Stop();
                FlushBuffer();
            };

            InitializeTray();
            RegisterHotkeys();
            HookKeyboard();

            // Kuchsiz (4GB RAM) kompyuterlarda klaviatura uzilib qolmasligi uchun Watchdog taymer
            watchdogTimer = new System.Windows.Forms.Timer();
            watchdogTimer.Interval = 10000; // Har 10 soniyada tekshirish
            watchdogTimer.Tick += (s, e) =>
            {
                try
                {
                    // Agar klaviaturada 3 soniyadan beri hech narsa yozilmagan bo'lsa, hookni yangilab qo'yish
                    if (isActive && (DateTime.Now - lastKeyTime).TotalSeconds >= 3)
                    {
                        UnhookKeyboard();
                        HookKeyboard();
                    }
                    // Operativ xotirani doim minimal (3-5 MB) darajada saqlash
                    SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, -1, -1);
                }
                catch { }
            };
            watchdogTimer.Start();

            // Agar litsenziya bo'lmasa xabar berish
            if (!isLicensedCached)
            {
                trayIcon.ShowBalloonTip(3000, "Alifbo Yordamchisi", "Dastur faollashtirilmagan. Faollashtirish uchun belgi ustiga 2 marta bosing.", ToolTipIcon.Warning);
            }

            // GitHub orqali avtomatik yangilanishni fonda tekshirish (15 soniyadan keyin)
            System.Windows.Forms.Timer initialTimer = new System.Windows.Forms.Timer();
            initialTimer.Interval = 15000;
            initialTimer.Tick += (s, e) =>
            {
                initialTimer.Stop();
                initialTimer.Dispose();
                CheckForUpdatesAsync(false);
            };
            initialTimer.Start();

            // Har 4 soatda tekshirish
            updateCheckTimer = new System.Windows.Forms.Timer();
            updateCheckTimer.Interval = 4 * 60 * 60 * 1000;
            updateCheckTimer.Tick += (s, e) => { CheckForUpdatesAsync(false); };
            updateCheckTimer.Start();
        }

        #region Keyboard Hook va Yozishda Avtomatik O'tkazish

        private void HookKeyboard()
        {
            hookProc = HookCallback;
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                hookId = SetWindowsHookEx(WH_KEYBOARD_LL, hookProc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private void UnhookKeyboard()
        {
            if (bufferTimer != null)
            {
                try
                {
                    bufferTimer.Stop();
                    FlushBuffer();
                }
                catch { }
            }

            if (hookId != IntPtr.Zero)
            {
                UnhookWindowsHookEx(hookId);
                hookId = IntPtr.Zero;
            }
        }

        private void SendDirectChar(char c)
        {
            try
            {
                INPUT[] inp = new INPUT[2];
                inp[0].type = INPUT_KEYBOARD;
                inp[0].U.ki.wVk = 0;
                inp[0].U.ki.wScan = (ushort)c;
                inp[0].U.ki.dwFlags = KEYEVENTF_UNICODE;
                inp[0].U.ki.dwExtraInfo = MAGIC_EXTRA;

                inp[1].type = INPUT_KEYBOARD;
                inp[1].U.ki.wVk = 0;
                inp[1].U.ki.wScan = (ushort)c;
                inp[1].U.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
                inp[1].U.ki.dwExtraInfo = MAGIC_EXTRA;

                uint sent = SendInput(2, inp, Marshal.SizeOf(typeof(INPUT)));
                if (sent == 0)
                {
                    try { SendKeys.SendWait(c.ToString()); } catch { }
                }
            }
            catch { }
        }

        private void FlushBuffer()
        {
            try
            {
                if (bufferedChar != '\0')
                {
                    char c = bufferedChar;
                    bufferedChar = '\0';
                    lastChar = c;
                    lastCharTime = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;
                    SendDirectChar(c);
                }
            }
            catch { }
        }

        private void SendReplacement(char replaceWith)
        {
            try
            {
                // 1. Shift (chap yoki o'ng) bosilganligini aniqlash
                bool lShift = (GetAsyncKeyState(0xA0) & 0x8000) != 0; // VK_LSHIFT
                bool rShift = (GetAsyncKeyState(0xA1) & 0x8000) != 0; // VK_RSHIFT
                bool anyShift = lShift || rShift || ((GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0) || ((GetKeyState(VK_SHIFT) & 0x8000) != 0);

                if (anyShift)
                {
                    keybd_event(0xA0, 0x2A, KEYEVENTF_KEYUP, (UIntPtr)0x414C4642); // Left Shift UP
                    keybd_event(0xA1, 0x36, KEYEVENTF_KEYUP, (UIntPtr)0x414C4642); // Right Shift UP
                    keybd_event((byte)VK_SHIFT, 0x2A, KEYEVENTF_KEYUP, (UIntPtr)0x414C4642);
                }

                // 2. Oldingi harfni o'chirish (Backspace)
                // Apparat darajasidagi haqiqiy scan code (0x0E) bilan yuboriladi
                INPUT[] bInputs = new INPUT[2];
                bInputs[0].type = INPUT_KEYBOARD;
                bInputs[0].U.ki.wVk = (ushort)VK_BACK;
                bInputs[0].U.ki.wScan = 0x0E; // Apparat Backspace scan-kodi!
                bInputs[0].U.ki.dwFlags = 0;
                bInputs[0].U.ki.dwExtraInfo = MAGIC_EXTRA;

                bInputs[1].type = INPUT_KEYBOARD;
                bInputs[1].U.ki.wVk = (ushort)VK_BACK;
                bInputs[1].U.ki.wScan = 0x0E;
                bInputs[1].U.ki.dwFlags = KEYEVENTF_KEYUP;
                bInputs[1].U.ki.dwExtraInfo = MAGIC_EXTRA;

                uint bSent = SendInput(2, bInputs, Marshal.SizeOf(typeof(INPUT)));
                if (bSent == 0)
                {
                    // Zaxira usul: agar SendInput bloklansa, to'g'ri scan-kod bilan keybd_event orqali o'chirish
                    keybd_event((byte)VK_BACK, 0x0E, 0, (UIntPtr)0x414C4642);
                    keybd_event((byte)VK_BACK, 0x0E, KEYEVENTF_KEYUP, (UIntPtr)0x414C4642);
                }

                // 3. Matn muharriri (Word/Chrome/Telegram) oldingi harfni to'liq o'chirib ulgurishi uchun pauza
                Thread.Sleep(20);

                // 4. Yangi Unicode harfni (ö, ş, ç, ğ, Ö, Ş, Ç, Ğ) kiritish
                INPUT[] inpChar = new INPUT[2];
                inpChar[0].type = INPUT_KEYBOARD;
                inpChar[0].U.ki.wVk = 0; // MSDN: KEYEVENTF_UNICODE bilan wVk 0 bo'lishi shart!
                inpChar[0].U.ki.wScan = (ushort)replaceWith;
                inpChar[0].U.ki.dwFlags = KEYEVENTF_UNICODE;
                inpChar[0].U.ki.dwExtraInfo = MAGIC_EXTRA;

                inpChar[1].type = INPUT_KEYBOARD;
                inpChar[1].U.ki.wVk = 0;
                inpChar[1].U.ki.wScan = (ushort)replaceWith;
                inpChar[1].U.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
                inpChar[1].U.ki.dwExtraInfo = MAGIC_EXTRA;

                uint sent = SendInput(2, inpChar, Marshal.SizeOf(typeof(INPUT)));
                if (sent == 0)
                {
                    try { SendKeys.SendWait(replaceWith.ToString()); } catch { }
                }

                // 5. Agar foydalanuvchi barmoq bilan Shift ni bosib turgan bo'lsa, Shift holatini qayta tiklaymiz
                if (anyShift)
                {
                    if (lShift) keybd_event(0xA0, 0x2A, 0, (UIntPtr)0x414C4642);
                    if (rShift) keybd_event(0xA1, 0x36, 0, (UIntPtr)0x414C4642);
                    if (!lShift && !rShift) keybd_event((byte)VK_SHIFT, 0x2A, 0, (UIntPtr)0x414C4642);
                }
            }
            catch { }
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                lastKeyTime = DateTime.Now;
            }

            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN && isActive && isLicensedCached)
            {
                KBDLLHOOKSTRUCT kbd = (KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(KBDLLHOOKSTRUCT));

                // Faqat o'zimiz yuborgan signallarni o'tkazib yuboramiz
                // (AnyDesk, TeamViewer va boshqa masofaviy dasturlarda ham 100% ishlashi uchun)
                if (kbd.dwExtraInfo == MAGIC_EXTRA)
                {
                    return CallNextHookEx(hookId, nCode, wParam, lParam);
                }

                // Ctrl yoki Alt bosilgan bo'lsa (masalan Ctrl+C, Alt+Tab, Ctrl+S) buferni va xotirani tozalaymiz
                bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
                bool alt = (GetKeyState(VK_MENU) & 0x8000) != 0;
                if (ctrl || alt)
                {
                    if (useBufferMode && bufferTimer != null)
                    {
                        bufferTimer.Stop();
                        FlushBuffer();
                    }
                    lastChar = '\0';
                    return CallNextHookEx(hookId, nCode, wParam, lParam);
                }

                // Shift tugmasining o'zi yolg'iz bosilganda (VK_SHIFT, VK_LSHIFT, VK_RSHIFT)
                // Buferga tegilmaydi (chunki Sh yoki O' yozilayotgan bo'lishi mumkin)
                if (kbd.vkCode == VK_SHIFT || kbd.vkCode == 0xA0 || kbd.vkCode == 0xA1)
                {
                    return CallNextHookEx(hookId, nCode, wParam, lParam);
                }

                // Backspace bosilganda
                if (kbd.vkCode == VK_BACK)
                {
                    if (useBufferMode && bufferTimer != null)
                    {
                        bufferTimer.Stop();
                        if (bufferedChar != '\0')
                        {
                            // Agar harf hali ekranga chiqmasdan buferda turgan bo'lsa,
                            // uni shunchaki bekor qilamiz va Backspace ni bloklaymiz (ekranda hech narsa o'chishi shart emas)
                            bufferedChar = '\0';
                            lastChar = '\0';
                            return (IntPtr)1;
                        }
                    }
                    lastChar = '\0';
                    return CallNextHookEx(hookId, nCode, wParam, lParam);
                }

                // Bo'shliq, Enter, Tab yoki Escape bosilsa buferni to'kib, o'tkazib yuboramiz
                if (kbd.vkCode == VK_SPACE || kbd.vkCode == VK_RETURN || kbd.vkCode == VK_TAB || kbd.vkCode == VK_ESCAPE)
                {
                    if (useBufferMode && bufferTimer != null)
                    {
                        bufferTimer.Stop();
                        FlushBuffer();
                    }
                    lastChar = '\0';
                    return CallNextHookEx(hookId, nCode, wParam, lParam);
                }

                bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
                bool caps = (GetKeyState(VK_CAPITAL) & 0x0001) != 0;
                bool isUpper = shift ^ caps;

                char keyChar = '\0';

                // Harflar A-Z
                if (kbd.vkCode >= 0x41 && kbd.vkCode <= 0x5A)
                {
                    keyChar = (char)(isUpper ? ('A' + (kbd.vkCode - 0x41)) : ('a' + (kbd.vkCode - 0x41)));
                }
                // Apostrof: VK_OEM_7 (0xDE) yoki VK_OEM_3 (0xC0 - backtick `)
                // O' yoki G' yozayotganda barmoq Shift da qolib ketgan bo'lsa ham apostrof deb taniydi
                else if (kbd.vkCode == VK_OEM_7 || kbd.vkCode == VK_OEM_3)
                {
                    keyChar = '\'';
                }

                // --- 1. TEZKOR BUFER REJIMI (Backspace ishlatilmaydi, sş, oö bo'lish xavfi 0%) ---
                if (useBufferMode)
                {
                    // A) Agar buferda oldingi harf (s, c, o, g) kutib turgan bo'lsa:
                    if (bufferedChar != '\0')
                    {
                        char replaceWith = '\0';

                        // Juftlikni tekshirish
                        if ((bufferedChar == 'o' || bufferedChar == 'O') && keyChar == '\'')
                        {
                            replaceWith = (bufferedChar == 'o') ? '\u00F6' : '\u00D6'; // ö / Ö
                        }
                        else if ((bufferedChar == 'g' || bufferedChar == 'G') && keyChar == '\'')
                        {
                            replaceWith = (bufferedChar == 'g') ? '\u011F' : '\u011E'; // ğ / Ğ
                        }
                        else if ((bufferedChar == 's' || bufferedChar == 'S') && (keyChar == 'h' || keyChar == 'H'))
                        {
                            replaceWith = (bufferedChar == 's') ? '\u015F' : '\u015E'; // ş / Ş
                        }
                        else if ((bufferedChar == 'c' || bufferedChar == 'C') && (keyChar == 'h' || keyChar == 'H'))
                        {
                            replaceWith = (bufferedChar == 'c') ? '\u00E7' : '\u00C7'; // ç / Ç
                        }

                        // Agar juftlik mos kelsa: TO'G'RIDAN-TO'G'RI yangi harfni chiqaramiz (Backspace KERAK EMAS!)
                        if (replaceWith != '\0')
                        {
                            bufferTimer.Stop();
                            bufferedChar = '\0';
                            lastChar = '\0';
                            lastCharTime = 0;

                            SendDirectChar(replaceWith);
                            return (IntPtr)1; // h yoki ' tugmasini bloklash
                        }

                        // Juftlik mos kelmasa: avval buferdagi harfni chiqaramiz
                        bufferTimer.Stop();
                        char toFlush = bufferedChar;
                        bufferedChar = '\0';
                        lastChar = toFlush;
                        lastCharTime = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;

                        // Agar yangi bosilgan harf ham s, c, o, g bo'lsa (masalan "so", "os"), uni buferga olamiz
                        bool isNewCandidate = (keyChar == 's' || keyChar == 'S' ||
                                               keyChar == 'c' || keyChar == 'C' ||
                                               keyChar == 'o' || keyChar == 'O' ||
                                               keyChar == 'g' || keyChar == 'G');
                        if (isNewCandidate)
                        {
                            SendDirectChar(toFlush);
                            bufferedChar = keyChar;
                            bufferTimer.Start();
                            return (IntPtr)1;
                        }

                        // Agar oddiy harf bo'lsa (masalan "salom" dagi 'a'): ikkala harfni bir vaqtda chiqaramiz
                        if (keyChar != '\0')
                        {
                            SendDirectChar(toFlush);
                            SendDirectChar(keyChar);
                            lastChar = keyChar;
                            lastCharTime = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;
                            return (IntPtr)1;
                        }
                        else
                        {
                            // Boshqa harf bo'lmagan tugmalar (strelkalar va h.k.)
                            SendDirectChar(toFlush);
                            return CallNextHookEx(hookId, nCode, wParam, lParam);
                        }
                    }

                    // B) Agar bufer bo'sh bo'lsa va s, c, o, g bosilgan bo'lsa -> buferga olamiz (200 ms kutish)
                    bool isCandidate = (keyChar == 's' || keyChar == 'S' ||
                                        keyChar == 'c' || keyChar == 'C' ||
                                        keyChar == 'o' || keyChar == 'O' ||
                                        keyChar == 'g' || keyChar == 'G');
                    if (isCandidate)
                    {
                        bufferedChar = keyChar;
                        bufferTimer.Stop();
                        bufferTimer.Start();
                        return (IntPtr)1; // Harfni ushlab turamiz, ekranga chiqarmaymiz
                    }
                }

                // --- 2. ZAXIRA / KLASSIK REJIM (Juda sekin terilganda yoki bufer o'chirilganda) ---
                if (keyChar != '\0')
                {
                    long now = DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond;
                    char replaceWith = '\0';

                    // Agar oldingi harf bilan oraliq 1.5 soniyadan kam bo'lsa
                    if (now - lastCharTime < 1500 && lastChar != '\0')
                    {
                        // o' -> ö, O' -> Ö
                        if ((lastChar == 'o' || lastChar == 'O') && keyChar == '\'')
                        {
                            replaceWith = (lastChar == 'o') ? '\u00F6' : '\u00D6'; // ö / Ö
                        }
                        // g' -> ğ, G' -> Ğ
                        else if ((lastChar == 'g' || lastChar == 'G') && keyChar == '\'')
                        {
                            replaceWith = (lastChar == 'g') ? '\u011F' : '\u011E'; // ğ / Ğ
                        }
                        // sh -> ş, Sh/SH -> Ş, sH -> ş
                        else if ((lastChar == 's' || lastChar == 'S') && (keyChar == 'h' || keyChar == 'H'))
                        {
                            replaceWith = (lastChar == 's') ? '\u015F' : '\u015E'; // ş / Ş
                        }
                        // ch -> ç, Ch/CH -> Ç, cH -> ç
                        else if ((lastChar == 'c' || lastChar == 'C') && (keyChar == 'h' || keyChar == 'H'))
                        {
                            replaceWith = (lastChar == 'c') ? '\u00E7' : '\u00C7'; // ç / Ç
                        }
                    }

                    if (replaceWith != '\0')
                    {
                        // O'RNATISH:
                        // 1. lastChar xotirani darhol tozalash (takrorlanish bo'lmasligi uchun)
                        lastChar = '\0';
                        lastCharTime = 0;

                        // 2. Backspace va yangi harfni jo'natish
                        SendReplacement(replaceWith);

                        // 3. Joriy tugma (' yoki h) dasturga o'tmasligi uchun bloklash
                        return (IntPtr)1;
                    }

                    // Aks holda joriy harfni eslab qolamiz
                    lastChar = keyChar;
                    lastCharTime = now;
                }
                else
                {
                    // Boshqa belgilar terilganda xotirani tozalash
                    lastChar = '\0';
                }
            }

            return CallNextHookEx(hookId, nCode, wParam, lParam);
        }

        #endregion

        #region Litsenziyalash Tizimi (Hardware ID)

        private static int Hash16(string s)
        {
            int h = 17;
            for (int i = 0; i < s.Length; i++)
            {
                int c = (int)s[i];
                h = ((h * 31 + c) & 32767);
            }
            return h;
        }

        public static string GetMachineCode()
        {
            string sn = "12345678";
            try
            {
                Type fsoType = Type.GetTypeFromProgID("Scripting.FileSystemObject");
                if (fsoType != null)
                {
                    dynamic fso = Activator.CreateInstance(fsoType);
                    dynamic drv = fso.GetDrive("C:");
                    sn = Math.Abs((long)drv.SerialNumber).ToString();
                }
            }
            catch { }

            int h1 = Hash16("M1" + sn);
            int h2 = Hash16("M2" + sn);
            return string.Format("AF-{0:X4}-{1:X4}", h1, h2);
        }

        public static string CalculateKey(string machCode)
        {
            string clean = machCode.Replace("-", "").Replace(" ", "").ToUpper();
            string s = SALT + clean;
            int k1 = Hash16("K1" + s);
            int k2 = Hash16("K2" + s);
            return string.Format("AK-{0:X4}-{1:X4}", k1, k2);
        }

        public static bool IsLicensed()
        {
            string mc = GetMachineCode();
            string validKey = CalculateKey(mc);

            string saved = "";
            try
            {
                using (RegistryKey rk = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\AlifboSetup"))
                {
                    if (rk != null)
                    {
                        object val = rk.GetValue("lic_key");
                        if (val != null) saved = val.ToString().Trim();
                    }
                }
            }
            catch { }

            if (string.IsNullOrEmpty(saved))
            {
                try
                {
                    string dat = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Alifbo\lic.dat");
                    if (File.Exists(dat))
                    {
                        saved = File.ReadAllText(dat).Trim();
                    }
                }
                catch { }
            }

            return !string.IsNullOrEmpty(saved) && string.Equals(saved, validKey, StringComparison.OrdinalIgnoreCase);
        }

        public static void SaveLicenseKey(string key)
        {
            try
            {
                using (RegistryKey rk = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\AlifboSetup"))
                {
                    if (rk != null)
                    {
                        rk.SetValue("lic_key", key);
                    }
                }
            }
            catch { }

            try
            {
                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Alifbo");
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                File.WriteAllText(Path.Combine(dir, "lic.dat"), key);
            }
            catch { }
        }

        private void CheckOrAskLicense()
        {
            string mc = GetMachineCode();
            string valid = CalculateKey(mc);

            isLicensedCached = IsLicensed();

            if (isLicensedCached)
            {
                MessageBox.Show(
                    string.Format("Litsenziya faol!\n\nKompyuter kodi: {0}\nLitsenziya kaliti: {1}\n\nMuallif: Fozliddinov Safarmurod\nTelegram: @Safar_master\nTelefon: +998939530473", mc, valid),
                    "AlifboO'tkazgich Litsenziyasi",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            using (Form f = new Form())
            {
                f.Text = "AlifboO'tkazgich: Litsenziya";
                f.Size = new Size(420, 260);
                f.StartPosition = FormStartPosition.CenterScreen;
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.MaximizeBox = false;
                f.MinimizeBox = false;

                Label lbl1 = new Label() { Text = "Sizning kompyuter kodingiz:", Location = new Point(20, 15), Size = new Size(360, 20) };
                TextBox txtMc = new TextBox() { Text = mc, Location = new Point(20, 40), Size = new Size(360, 24), ReadOnly = true };

                Label lbl2 = new Label() { Text = "Litsenziya kalitini kiriting (Muallif: @Safar_master):", Location = new Point(20, 75), Size = new Size(360, 20) };
                TextBox txtKey = new TextBox() { Location = new Point(20, 100), Size = new Size(360, 24) };

                Button btnOk = new Button() { Text = "Faollashtirish", Location = new Point(140, 150), Size = new Size(120, 32), DialogResult = DialogResult.OK };
                Button btnCancel = new Button() { Text = "Bekor qilish", Location = new Point(270, 150), Size = new Size(110, 32), DialogResult = DialogResult.Cancel };

                f.Controls.AddRange(new Control[] { lbl1, txtMc, lbl2, txtKey, btnOk, btnCancel });
                f.AcceptButton = btnOk;
                f.CancelButton = btnCancel;

                if (f.ShowDialog() == DialogResult.OK)
                {
                    string input = txtKey.Text.Trim();
                    if (string.Equals(input, valid, StringComparison.OrdinalIgnoreCase))
                    {
                        SaveLicenseKey(valid);
                        isLicensedCached = true;
                        UpdateTrayStatus();
                        MessageBox.Show("Litsenziya muvaffaqiyatli faollashtirildi! Rahmat.", "AlifboO'tkazgich", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Litsenziya kaliti noto'g'ri!\n\nMuallifga murojaat qiling: @Safar_master (+998939530473)", "Xatolik", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        #endregion

        #region GitHub Auto-Update Tizimi

        public void CheckForUpdatesAsync(bool showUpToDateMessage)
        {
            ThreadPool.QueueUserWorkItem((state) =>
            {
                try
                {
                    string json = "";
                    using (WebClient wc = new WebClient())
                    {
                        wc.Headers.Add("User-Agent", "AlifboOtkazgich-Updater");
                        json = wc.DownloadString(GITHUB_VERSION_URL);
                    }

                    int idx = json.IndexOf("\"version\"");
                    if (idx >= 0)
                    {
                        int colon = json.IndexOf(':', idx);
                        int q1 = json.IndexOf('"', colon + 1);
                        int q2 = json.IndexOf('"', q1 + 1);
                        string latestVer = json.Substring(q1 + 1, q2 - q1 - 1).Trim();

                        if (!string.IsNullOrEmpty(latestVer) && string.Compare(latestVer, CURRENT_VERSION, StringComparison.OrdinalIgnoreCase) > 0)
                        {
                            DialogResult dr = MessageBox.Show(
                                string.Format("AlifboO'tkazgichning yangi versiyasi ({0}) mavjud!\n\nJoriy versiya: {1}\n\nYangilashni hozir yuklab olib o'rnataylikmi?", latestVer, CURRENT_VERSION),
                                "AlifboO'tkazgich yangilanishi",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Information);

                            if (dr == DialogResult.Yes)
                            {
                                DownloadAndApplyUpdate();
                            }
                        }
                        else if (showUpToDateMessage)
                        {
                            MessageBox.Show("Sizda eng so'nggi versiya o'rnatilgan (" + CURRENT_VERSION + ").", "Alifbo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (showUpToDateMessage)
                    {
                        MessageBox.Show("Yangilanishlarni tekshirishda xatolik yuz berdi:\n" + ex.Message, "Alifbo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            });
        }

        private void DownloadAndApplyUpdate()
        {
            try
            {
                string tempSetup = Path.Combine(Path.GetTempPath(), "AlifboOtkazgich-Update.exe");
                using (WebClient wc = new WebClient())
                {
                    wc.Headers.Add("User-Agent", "AlifboOtkazgich-Updater");
                    wc.DownloadFile(GITHUB_SETUP_URL, tempSetup);
                }

                if (File.Exists(tempSetup))
                {
                    Process.Start(tempSetup, "/S");
                    Application.Exit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Yangilanishni yuklab olishda xatolik: " + ex.Message, "Alifbo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        private void InitializeTray()
        {
            trayMenu = new ContextMenuStrip();

            itemActive = new ToolStripMenuItem("Faol (Ishlamoqda)", null, ToggleActive);
            itemActive.Checked = true;
            itemActive.Font = new Font(itemActive.Font, FontStyle.Bold);
            trayMenu.Items.Add(itemActive);

            ToolStripMenuItem itemBuffer = new ToolStripMenuItem("Tezkor bufer rejimi (Tavsiya)", null, (s, e) =>
            {
                useBufferMode = !useBufferMode;
                ((ToolStripMenuItem)s).Checked = useBufferMode;
                if (!useBufferMode && bufferTimer != null)
                {
                    bufferTimer.Stop();
                    FlushBuffer();
                }
            });
            itemBuffer.Checked = true;
            trayMenu.Items.Add(itemBuffer);

            trayMenu.Items.Add(new ToolStripSeparator());

            trayMenu.Items.Add("Ctrl+Shift+V: O'tkazib joylash");
            trayMenu.Items.Add("Ctrl+Alt+U: Yoqish / O'chirish");
            trayMenu.Items.Add(new ToolStripSeparator());

            trayMenu.Items.Add("Yangilanishlarni tekshirish...", null, (s, e) => CheckForUpdatesAsync(true));
            trayMenu.Items.Add("Litsenziya", null, (s, e) => CheckOrAskLicense());
            trayMenu.Items.Add("Dastur haqida", null, (s, e) => ShowAbout());
            trayMenu.Items.Add(new ToolStripSeparator());

            trayMenu.Items.Add("Chiqish", null, Exit);

            trayIcon = new NotifyIcon();
            trayIcon.Text = "Alifbo Yordamchisi (Ctrl+Alt+U)";

            try
            {
                string iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "alifbo.ico");
                if (File.Exists(iconPath))
                {
                    trayIcon.Icon = new Icon(iconPath);
                }
                else
                {
                    trayIcon.Icon = SystemIcons.Application;
                }
            }
            catch
            {
                trayIcon.Icon = SystemIcons.Application;
            }

            trayIcon.ContextMenuStrip = trayMenu;
            trayIcon.Visible = true;
            trayIcon.DoubleClick += (s, e) =>
            {
                if (!isLicensedCached)
                {
                    CheckOrAskLicense();
                }
                else
                {
                    ToggleActive(s, e);
                }
            };

            UpdateTrayStatus();
        }

        private void UpdateTrayStatus()
        {
            if (!isLicensedCached)
            {
                trayIcon.Text = "Alifbo Yordamchisi (Faollashtirilmagan!)";
                itemActive.Text = "Faollashtirilmagan (Litsenziya kerak)";
                itemActive.Checked = false;
            }
            else
            {
                itemActive.Text = isActive ? "Faol (Ishlamoqda)" : "To'xtatilgan";
                itemActive.Checked = isActive;
                trayIcon.Text = isActive ? "Alifbo Yordamchisi (Faol)" : "Alifbo Yordamchisi (To'xtatilgan)";
            }
        }

        private void ToggleActive(object sender, EventArgs e)
        {
            if (!isLicensedCached)
            {
                CheckOrAskLicense();
                return;
            }

            isActive = !isActive;
            UpdateTrayStatus();

            string msg = isActive ? "Alifbo Yordamchisi yoqildi" : "Alifbo Yordamchisi to'xtatildi";
            trayIcon.ShowBalloonTip(1500, "Alifbo Yordamchisi", msg, ToolTipIcon.Info);
        }

        private void RegisterHotkeys()
        {
            hiddenWindow = new HiddenWindow(this);
            RegisterHotKey(hiddenWindow.Handle, HOTKEY_PASTE, MOD_CONTROL | MOD_SHIFT, 0x56); // Ctrl+Shift+V
            RegisterHotKey(hiddenWindow.Handle, HOTKEY_TOGGLE, MOD_CONTROL | MOD_ALT, 0x55); // Ctrl+Alt+U
        }

        public void HandleHotkey(int id)
        {
            if (id == HOTKEY_TOGGLE)
            {
                ToggleActive(null, null);
            }
            else if (id == HOTKEY_PASTE)
            {
                if (!isActive || !isLicensedCached) return;
                PasteConvertedText();
            }
        }

        public static string ConvertToNewAlphabet(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            const string APOS = "'\u2018\u2019\u201B\u02B9\u02BB\u02BC\u02C8\u02CA\u2032`\u00B4\uFF07";
            StringBuilder sb = new StringBuilder();
            int n = text.Length;

            for (int i = 0; i < n; i++)
            {
                char c = text[i];
                char c1 = (i + 1 < n) ? text[i + 1] : '\0';

                // o' -> ö, O' -> Ö
                if ((c == 'o' || c == 'O') && c1 != '\0' && APOS.IndexOf(c1) >= 0)
                {
                    sb.Append(c == 'o' ? '\u00F6' : '\u00D6');
                    i++;
                }
                // g' -> ğ, G' -> Ğ
                else if ((c == 'g' || c == 'G') && c1 != '\0' && APOS.IndexOf(c1) >= 0)
                {
                    sb.Append(c == 'g' ? '\u011F' : '\u011E');
                    i++;
                }
                // sh -> ş, Sh/SH -> Ş
                else if ((c == 's' || c == 'S') && (c1 == 'h' || c1 == 'H'))
                {
                    sb.Append(c == 's' ? '\u015F' : '\u015E');
                    i++;
                }
                // ch -> ç, Ch/CH -> Ç
                else if ((c == 'c' || c == 'C') && (c1 == 'h' || c1 == 'H'))
                {
                    sb.Append(c == 'c' ? '\u00E7' : '\u00C7');
                    i++;
                }
                else
                {
                    sb.Append(c);
                }
            }

            return sb.ToString();
        }

        private void PasteConvertedText()
        {
            try
            {
                if (!Clipboard.ContainsText()) return;
                string original = Clipboard.GetText();
                string converted = ConvertToNewAlphabet(original);

                Clipboard.SetText(converted);
                SendKeys.SendWait("^v");
            }
            catch { }
        }

        private void ShowAbout()
        {
            MessageBox.Show(
                string.Format("AlifboO'tkazgich Yordamchisi v{0}\n\nWord, Excel, PowerPoint va butun Windows uchun yangi o'zbek alifbosi yordamchisi.\n\nImkoniyatlar:\n• Yozish jarayonida avtomatik o'tkazish (o' -> ö, sh -> ş, ch -> ç, g' -> ğ)\n• Barcha Office dasturlarida 1 tugma bilan to'liq o'tkazish\n• Ctrl+Shift+V: O'tkazib joylash\n• Ctrl+Alt+U: Yordamchini yoqish/o'chirish\n• GitHub orqali avtomatik yangilanish\n\nMuallif: Fozliddinov Safarmurod\nTelegram: @Safar_master\nTelefon: +998939530473", CURRENT_VERSION),
                "AlifboO'tkazgich haqida",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void Exit(object sender, EventArgs e)
        {
            UnhookKeyboard();

            if (watchdogTimer != null)
            {
                watchdogTimer.Stop();
                watchdogTimer.Dispose();
            }

            if (updateCheckTimer != null)
            {
                updateCheckTimer.Stop();
                updateCheckTimer.Dispose();
            }

            if (hiddenWindow != null)
            {
                UnregisterHotKey(hiddenWindow.Handle, HOTKEY_PASTE);
                UnregisterHotKey(hiddenWindow.Handle, HOTKEY_TOGGLE);
                hiddenWindow.Dispose();
            }

            trayIcon.Visible = false;
            Application.Exit();
        }

        private class HiddenWindow : Form
        {
            private TrayApplicationContext context;

            public HiddenWindow(TrayApplicationContext context)
            {
                this.context = context;
                this.WindowState = FormWindowState.Minimized;
                this.ShowInTaskbar = false;
                this.FormBorderStyle = FormBorderStyle.None;
                this.Size = new Size(0, 0);
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WM_HOTKEY)
                {
                    context.HandleHotkey(m.WParam.ToInt32());
                }
                base.WndProc(ref m);
            }
        }
    }
}
