using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using WindowsInput;
using WindowsInput.Native;

namespace FasType.LLKeyboardListener
{
    public static class Input
    {
        static readonly InputSimulator _sim;
        static IKeyboardSimulator Keyboard => _sim.Keyboard;

        static Input() => _sim = new();



        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        public static void Erase(int n)
        {
            //var activeHwnd = GetFocus();
            //if (activeHwnd != Utils.Caret.CurrentCaretHwnd)
            SetForegroundWindow(Utils.Caret.CurrentCaretHwnd);
            //if (SetForegroundWindow(Utils.Caret.CurrentCaretHwnd) == false)
            //{
            //    int err = Marshal.GetLastWin32Error();
            //}

            Keyboard.KeyPress(Enumerable.Repeat(VirtualKeyCode.BACK, n).ToArray());
        }

        public static void TextEntry(string text)
        {
            //var activeHwnd = GetFocus();
            //if (activeHwnd != Utils.Caret.CurrentCaretHwnd)
            SetForegroundWindow(Utils.Caret.CurrentCaretHwnd);
            //if (SetForegroundWindow(Utils.Caret.CurrentCaretHwnd) == false)
            //{
            //    int err = Marshal.GetLastWin32Error();
            //}

            Keyboard.TextEntry(text);
        }

        //public static void KeyPress(VirtualKeyCode keyCode) => Keyboard.KeyPress(keyCode);
        //public static void KeyPress(params VirtualKeyCode[] keyCodes) => Keyboard.KeyPress(keyCodes);
        //public static void TextEntry(char character) => Keyboard.TextEntry(character);
    }
}
