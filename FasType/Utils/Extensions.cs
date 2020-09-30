using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;

namespace FasType.Utils
{
    public static class KeyExtensions
    {
        public static bool IsAlpha(this Key k) => k is >= Key.A and <= Key.Z || (k is Key.D2 or Key.D7 or Key.D9 or Key.D0 or Key.Oem3 && !KeyboardStates.IsModified());
    }

    public static class StringExtensions
    {
        public static bool IsFirstCharUpper(this string input) => input switch
        {
            null => throw new ArgumentNullException(nameof(input)),
            "" => throw new ArgumentException($"{nameof(input)} cannot be empty", nameof(input)),
            _ => char.IsUpper(input[0])
        };

        public static string FirstCharToUpper(this string input) => input switch
        {
             null => throw new ArgumentNullException(nameof(input)),
             "" => throw new ArgumentException($"{nameof(input)} cannot be empty", nameof(input)),
             _ => input.First().ToString().ToUpper() + input.Substring(1)
        };
    }
}
