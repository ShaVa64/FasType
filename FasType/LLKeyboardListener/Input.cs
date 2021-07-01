using FasType.Utils;
using Microsoft.Extensions.DependencyInjection;
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
        private static readonly InputSimulator _sim;
        private static IKeyboardSimulator Keyboard => _sim.Keyboard;

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        static Input() => _sim = new();

        public static void Erase(int n)
        {
            SetForegroundWindow(Caret.CurrentCaretHwnd);
            Keyboard.KeyPress(Enumerable.Repeat(VirtualKeyCode.BACK, n).ToArray());
        }

        public static void TextEntry(string text)
        {
            SetForegroundWindow(Caret.CurrentCaretHwnd);
            Keyboard.TextEntry(text);
        }
    }
}
