using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Input;

namespace FasType.Utils
{
    public static class KeyboardStates
    {
        public static bool IsShifted() => Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift) || Keyboard.IsKeyToggled(Key.Capital);

        public static bool IsModified() => IsShifted() || Keyboard.IsKeyDown(Key.LeftAlt) || Keyboard.IsKeyDown(Key.RightAlt) || Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl);
    }
}
