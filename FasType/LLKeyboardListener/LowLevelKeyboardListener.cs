using FasType.Utils;
using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Input;

namespace FasType.LLKeyboardListener
{
    public class LowLevelKeyboardListener
    {
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_SYSKEYDOWN = 0x0104;

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);



        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);
        [StructLayout(LayoutKind.Sequential)] public struct GUITHREADINFO
        {
            public int cbSize;
            public int flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public System.Drawing.Rectangle rcCaret;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);
        [StructLayout(LayoutKind.Sequential)] public struct RECT
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }


        public delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        public event EventHandler<KeyPressedEventArgs>? OnKeyPressed;

        KeyPressedEventArgs.UniqueKeyPressed? oldKey, newKey;

        private readonly LowLevelKeyboardProc _proc;
        private IntPtr _hookID = IntPtr.Zero;

        public LowLevelKeyboardListener()
        {
            _proc = HookCallback;

            oldKey = null;
            newKey = null;
        }

        public void HookKeyboard()
        {
            if (_hookID == IntPtr.Zero)
            {
                _hookID = SetHook(_proc);
                Log.Information("{kblisetner} was hooked.", nameof(LowLevelKeyboardListener));
            }
            else
            {
                Log.Information("{kblisetner} was already hooked.", nameof(LowLevelKeyboardListener));
            }
        }

        public void UnHookKeyboard()
        {
            if (_hookID != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookID);
                _hookID = IntPtr.Zero;
                Log.Information("{kblisetner} was unhooked.", nameof(LowLevelKeyboardListener));
            }
            else
            {
                Log.Information("{kblisetner} was already unhooked.", nameof(LowLevelKeyboardListener));
            }
        }

        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using Process curProcess = Process.GetCurrentProcess();
            using ProcessModule? curModule = curProcess.MainModule;
            _ = curModule ?? throw new NullReferenceException();
            var mName = curModule.ModuleName ?? throw new NullReferenceException();
            return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(mName), 0);
        }

        private System.Windows.Point CaretPos()
        {
            int hwnd = 0;
            GUITHREADINFO guiti = new();
            guiti.cbSize = Marshal.SizeOf(guiti);

            var b1 = GetGUIThreadInfo(0, ref guiti);

            System.Windows.Point p = new(guiti.rcCaret.Left, guiti.rcCaret.Top);
            Debug.WriteLine($"Caret Inside: ({p.X}, {p.Y})");

            var b2 = GetWindowRect(guiti.hwndActive, out RECT rect);
            p.Offset(rect.Left, rect.Top);
            Debug.WriteLine($"Caret Outside: ({p.X}, {p.Y})");

            System.Drawing.Point dp = new((int)p.X, (int)p.Y);
            var screen = System.Windows.Forms.Screen.FromPoint(dp);
            var wa = System.Windows.Forms.Screen.GetWorkingArea(dp);

            Debug.WriteLine($"Caret Screen Name (Primary): {screen.DeviceName} ({screen.Primary})");
            Debug.WriteLine($"Working Area: ({wa.Left}, {wa.Top}) {wa.Width}x{wa.Height}");
            return p;
        }

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            KeyPressedEventArgs? eventArgs = null;
            CaretPos();
            if (OnKeyPressed is not null && nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN))
            {
                int vkCode = Marshal.ReadInt32(lParam);
                Key key = KeyInterop.KeyFromVirtualKey(vkCode);
                newKey = new(key, vkCode);
                eventArgs = new (oldKey, newKey);
                OnKeyPressed(this, eventArgs);
            }
            oldKey = newKey;

            IntPtr r = CallNextHookEx(_hookID, nCode, wParam, lParam);
            return eventArgs?.StopChain == true ? (IntPtr)1 : r;
        }
    }

    public class KeyPressedEventArgs : EventArgs
    {
        public bool StopChain { get; set; }
        public UniqueKeyPressed? Old { get; }
        public UniqueKeyPressed New { get; }
        public Key KeyPressed => New.KeyPressed;
        public bool IsShifted => New.IsShifted;

        public KeyPressedEventArgs(UniqueKeyPressed? old, UniqueKeyPressed @new)
        {
            StopChain = false;
            Old = old;
            New = @new;
        }

        public class UniqueKeyPressed
        {
            public int VkCode { get; private set; }
            public Key KeyPressed { get; private set; }
            public bool IsShifted { get; private set; }
            //public bool IsSystemKey { get; private set; }

            public UniqueKeyPressed(Key key, int vkCode/*, bool isSystemKey*/)
            {
                KeyPressed = key;
                VkCode = vkCode;
                IsShifted = KeyboardStates.IsShifted();
                //IsSystemKey = isSystemKey;
            }
        }
    }
}