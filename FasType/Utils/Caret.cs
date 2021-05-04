using System;
using System.Collections.Generic;
using System.Diagnostics;
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
            public System.Drawing.Rectangle rcCaret;
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

        public static System.Drawing.Point GetCaretPos()
        {
            //int hwnd = 0;
            GUITHREADINFO guiti = new();
            guiti.cbSize = Marshal.SizeOf(guiti);

            var b1 = GetGUIThreadInfo(0, ref guiti);

            var p = guiti.rcCaret.Location;// new(guiti.rcCaret.Left, guiti.rcCaret.Top);
            //p.Offset(guiti.rcCaret.Width, guiti.rcCaret.Height);
            Debug.WriteLine($"Caret Inside: ({p.X}, {p.Y})");

            var b2 = GetWindowRect(guiti.hwndActive, out RECT rect);
            p.Offset(rect.Left, rect.Top);
            Debug.WriteLine($"Caret Outside: ({p.X}, {p.Y})");

            //System.Drawing.Point dp = new((int)p.X, (int)p.Y);
            return p;
        }

        public static System.Drawing.Rectangle GetWorkingArea(System.Drawing.Point dp)
        {
            var screen = System.Windows.Forms.Screen.FromPoint(dp);
            var wa = System.Windows.Forms.Screen.GetWorkingArea(dp);

            Debug.WriteLine($"Caret Screen Name (Primary): {screen.DeviceName} ({screen.Primary})");
            Debug.WriteLine($"Working Area: ({wa.Left}, {wa.Top}) {wa.Width}x{wa.Height}");

            return wa;
        }

    }
}
