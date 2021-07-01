using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;

namespace FasType.Utils
{
    public static class KeyExtensions
    {
        public static bool IsAlpha(this Key k) => k is >= Key.A and <= Key.Z || (k is Key.D2 or Key.D7 or Key.D9 or Key.D0 or Key.Oem3 && !KeyboardStates.IsModified());
        public static bool IsModifier(this Key k) => k is Key.LeftCtrl or Key.RightCtrl or Key.LeftShift or Key.RightShift or Key.LWin or Key.RWin or Key.CapsLock;
    }

    public static class AbbreviationExtensions
    {
        public static Type GetModifyPageType(this Core.Models.Abbreviations.BaseAbbreviation ba) => ba switch
        {
            Core.Models.Abbreviations.SimpleAbbreviation => typeof(Pages.SimpleAbbreviationPage),
            Core.Models.Abbreviations.VerbAbbreviation => throw new NotImplementedException(),
            _ => throw new NotImplementedException(),
        };
    }

    public static class StringExtensions
    {
        public static bool IsFirstCharUpper(this string input) => input switch
        {
            null => throw new ArgumentNullException(nameof(input)),
            "" => false, //throw new ArgumentException($"{nameof(input)} cannot be empty", nameof(input)),
            _ => char.IsUpper(input[0])
        };

        public static string FirstCharToUpper(this string input) => input switch
        {
             null => throw new ArgumentNullException(nameof(input)),
             "" => "",// throw new ArgumentException($"{nameof(input)} cannot be empty", nameof(input)),
             _ => input.First().ToString().ToUpper() + input[1..]
        };
    }

    public static class BoolExtensions
    {
        public static void IfTrue(this bool b, Action trueAction)
        {
            if (b)
            {
                trueAction();
            }
        }
    }

    //public static class LinqExtensions
    //{
    //    public static IEnumerable<T> Cast<T>(this IEnumerable<AbbreviationMethodRecord> enumerable) where T : AbbreviationMethod => enumerable.Select(sar => (T)sar);
    //    public static IEnumerable<T> Cast<T>(this IEnumerable<AbbreviationMethod> enumerable) where T : AbbreviationMethodRecord => enumerable.Select(sa => (T)sa);

    //    public static IEnumerable<T> Cast<T>(this IEnumerable<GrammarTypeRecord> enumerable) where T : GrammarType => enumerable.Select(gtr => (T)gtr);
    //    public static IEnumerable<T> Cast<T>(this IEnumerable<GrammarType> enumerable) where T : GrammarTypeRecord => enumerable.Select(gt => (T)gt);
    //}

    public static class ComboBoxExtensions
    {
        public static void SetWidthFromItems(this ComboBox comboBox)
        {
            if (comboBox.Template.FindName("PART_Popup", comboBox) is Popup popup
                && popup.Child is FrameworkElement popupContent)
            {
                popupContent.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                // suggested in comments, original answer has a static value 19.0
                var emptySize = SystemParameters.VerticalScrollBarWidth + comboBox.Padding.Left + comboBox.Padding.Right;
                comboBox.Width = emptySize + popupContent.DesiredSize.Width;
            }
        }
    }

    public static class WindowExtensions
    {
        private const UInt32 FLASHW_STOP = 0; //Stop flashing. The system restores the window to its original state.       
        private const UInt32 FLASHW_CAPTION = 1; //Flash the window caption.        
        private const UInt32 FLASHW_TRAY = 2; //Flash the taskbar button.        
        private const UInt32 FLASHW_ALL = 3; //Flash both the window caption and taskbar button.        
        private const UInt32 FLASHW_TIMER = 4; //Flash continuously, until the FLASHW_STOP flag is set.        
        private const UInt32 FLASHW_TIMERNOFG = 12; //Flash continuously until the window comes to the foreground.  

        private const UInt32 FLASH_TIMEOUT = 0;

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct FLASHWINFO
        {
            public UInt32 cbSize; //The size of the structure in bytes.            
            public IntPtr hwnd; //A Handle to the Window to be Flashed. The window can be either opened or minimized.

            public UInt32 dwFlags; //The Flash Status.            
            public UInt32 uCount; // number of times to flash the window            
            public UInt32 dwTimeout; //The rate at which the Window is to be flashed, in milliseconds. If Zero, the function uses the default cursor blink rate.        
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

        public static bool FlashWindow(this Window win, UInt32 count = UInt32.MaxValue)
        {
            //Don't flash if the window is active            
            //if (win.IsActive) return;
            System.Windows.Interop.WindowInteropHelper h = new(win);
            //Serilog.Log.Information($"FW, Handle: {h.Handle}");
            FLASHWINFO info = new()
            {
                hwnd = h.Handle,
                dwFlags = FLASHW_TRAY | FLASHW_TIMER,
                uCount = count,
                dwTimeout = FLASH_TIMEOUT
            };

            info.cbSize = Convert.ToUInt32(System.Runtime.InteropServices.Marshal.SizeOf(info));
            return FlashWindowEx(ref info);
        }

        public static bool FlashWindowUntillFocus(this Window win)
        {
            //Don't flash if the window is active            
            //if (win.IsActive) return;
            System.Windows.Interop.WindowInteropHelper h = new(win);
            //Serilog.Log.Information($"FWUF, Handle: {h.Handle}");
            FLASHWINFO info = new()
            {
                hwnd = h.Handle,
                dwFlags = FLASHW_TRAY | FLASHW_TIMERNOFG,
                uCount = uint.MaxValue,
                dwTimeout = FLASH_TIMEOUT
            };

            info.cbSize = Convert.ToUInt32(System.Runtime.InteropServices.Marshal.SizeOf(info));
            return FlashWindowEx(ref info);
        }

        public static bool StopFlashingWindow(this Window win)
        {
            System.Windows.Interop.WindowInteropHelper h = new(win);
            //Serilog.Log.Information($"SFW, Handle: {h.Handle}");
            FLASHWINFO info = new()
            {
                hwnd = h.Handle,
                dwFlags = FLASHW_STOP,
                uCount = UInt32.MaxValue,
                dwTimeout = FLASH_TIMEOUT
            };
            info.cbSize = Convert.ToUInt32(System.Runtime.InteropServices.Marshal.SizeOf(info));
            return FlashWindowEx(ref info);
        }
    }

    public static class AppExtensions
    {
        public static bool FlashApp(this App app) => app.MainWindow.FlashWindow();
        public static bool StopFlashingApp(this App app) => app.MainWindow.StopFlashingWindow();
    }
}
