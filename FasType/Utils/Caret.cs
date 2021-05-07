using Serilog;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Utils
{
    public static class Caret
    {
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);
        [StructLayout(LayoutKind.Sequential)]
        public struct GUITHREADINFO
        {
            public int cbSize;
            public int flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public RECT rcCaret;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool ClientToScreen(IntPtr hwnd, ref Point lpRect);

        public static IntPtr CurrentCaretHwnd { get; private set; }

        public static Point GetCaretPos()
        {
            //int hwnd = 0;
            GUITHREADINFO guiti = new();
            guiti.cbSize = Marshal.SizeOf(guiti);

            var b1 = GetGUIThreadInfo(0, ref guiti);
            if (CurrentCaretHwnd != guiti.hwndCaret)
            {
                CurrentCaretHwnd = guiti.hwndCaret;
                //return GetCaretPos();
            }
            Point p = new(guiti.rcCaret.Left, guiti.rcCaret.Bottom + 5);
            //p.Offset(guiti.rcCaret.Width, guiti.rcCaret.Height);
            Log.Debug($"Caret Inside, (bool): ({p.X}, {p.Y}), ({b1})");

            var b2 = ClientToScreen(CurrentCaretHwnd, ref p);
            //var b2 = GetWindowRect(guiti.hwndActive, out RECT rect);
            //p.Offset(rect.Left, rect.Top);
            Log.Debug($"Caret Outside, (bool): ({p.X}, {p.Y}), ({b2})");

            //System.Drawing.Point dp = new((int)p.X, (int)p.Y);
            return p;
        }

        public static Rectangle GetWorkingArea(Point dp)
        {
            //var screen = System.Windows.Forms.Screen.FromPoint(dp);
            //Debug.WriteLine($"Caret Screen Name (Primary): {screen.DeviceName} ({screen.Primary})");

            var wa = System.Windows.Forms.Screen.GetWorkingArea(dp);
            Log.Debug($"Working Area: ({wa.Left}, {wa.Top}) {wa.Width}x{wa.Height}");

            return wa;
        }

    }
}
