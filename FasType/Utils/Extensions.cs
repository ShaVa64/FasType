﻿using System;
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
             _ => input.First().ToString().ToUpper() + input.Substring(1)
        };
    }

    public static class ComboBoxExtensions
    {
        public static void SetWidthFromItems(this ComboBox comboBox)
        {
            if (comboBox.Template.FindName("PART_Popup", comboBox) is Popup popup
                && popup.Child is FrameworkElement popupContent)
            {
                popupContent.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                // suggested in comments, original answer has a static value 19.0
                var emptySize = /*SystemParameters.VerticalScrollBarWidth +*/ comboBox.Padding.Left + comboBox.Padding.Right;
                comboBox.Width = emptySize + popupContent.DesiredSize.Width;
            }
        }
    }
}
