using System;
using System.Collections.Generic;
using System.Linq;
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

        public static void Erase(int n) => Keyboard.KeyPress(Enumerable.Repeat(VirtualKeyCode.BACK, n).ToArray());
        public static void TextEntry(string text) => Keyboard.TextEntry(text);

        //public static void KeyPress(VirtualKeyCode keyCode) => Keyboard.KeyPress(keyCode);
        //public static void KeyPress(params VirtualKeyCode[] keyCodes) => Keyboard.KeyPress(keyCodes);
        //public static void TextEntry(char character) => Keyboard.TextEntry(character);
    }
}
