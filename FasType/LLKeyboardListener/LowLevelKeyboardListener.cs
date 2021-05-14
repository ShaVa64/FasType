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
        private const int VK_PACKET = 0xE7;

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);


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
                Log.Information("{kblisetner} was hooked, (hook id: {_hookID}).", nameof(LowLevelKeyboardListener), _hookID);
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

        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            KeyPressedEventArgs? eventArgs = null;

            //Log.Debug("Code: {nCode}, wParam: {wParam}, lParam: {lParam}", nCode, wParam, lParam);
            if (OnKeyPressed is not null && nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN))
            {
                int vkCode = Marshal.ReadInt32(lParam);
                if (vkCode != VK_PACKET)
                {
                    Key key = KeyInterop.KeyFromVirtualKey(vkCode);
                    newKey = new(key, vkCode);
                    eventArgs = new (oldKey, newKey);
                    OnKeyPressed(this, eventArgs);
                    //Log.Debug("Chain Stopped on {key}: {StopChain}", eventArgs?.KeyPressed, eventArgs?.StopChain);
                }
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