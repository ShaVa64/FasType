using FasType.Models;
using FasType.Services;
using FasType.Utils;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Windows.Input;
using WindowsInput;

namespace FasType.LLKeyboardListener
{
    class KeyboardListenerHandler : IKeyboardListenerHandler
    {
        public string CurrentWord { get; private set; }
        Action<string> CurrentWordCallback { get; set; }

        readonly LowLevelKeyboardListener _listener;
        readonly InputSimulator _sim;
        ListenerStates _currentListenerState;
        readonly IDataStorage _storage;

        public KeyboardListenerHandler(IDataStorage storage)
        {
            _listener = new LowLevelKeyboardListener();
            _sim = new InputSimulator();
            _storage = storage;
            _currentListenerState = ListenerStates.Inserting;
        }

        bool TryWriteAbbreviation(IAbbreviation abbrev, string shortForm)
        {
            if (abbrev.TryGetFullForm(shortForm, out string fullForm))
            {
                string word = CurrentWord.IsFirstCharUpper() ? fullForm.FirstCharToUpper() : fullForm;
                _sim.Keyboard.KeyPress(Enumerable.Repeat(WindowsInput.Native.VirtualKeyCode.BACK, CurrentWord.Length).ToArray());
                _sim.Keyboard.TextEntry(word + " ");
                return true;
            }
            return false;
        }

        void ListenerEvent(object sender, KeyPressedArgs e)
        {
            Log.Information("Current Listener State: {listenerState}", _currentListenerState);
            if (_currentListenerState is ListenerStates.Inserting)
                Inserting(sender, e);
            else if (_currentListenerState is ListenerStates.Choosing)
                Choosing(sender, e);

            CurrentWordCallback?.Invoke(CurrentWord);
        }

        void Inserting(object sender, KeyPressedArgs e)
        {
            if (e.KeyPressed == Key.Space)
            {
                string shortForm = CurrentWord.ToLower();
        
                var abbrevs = _storage.GetAbbreviations(shortForm).ToList();

                if (abbrevs.Count == 1)
                {
                    var abbrev = abbrevs.Single();
                    e.StopChain |= TryWriteAbbreviation(abbrev, shortForm);
                }
                else if (abbrevs.Count > 1)
                {
                    foreach (var abbrev in abbrevs)
                    {
                        e.StopChain |= TryWriteAbbreviation(abbrev, shortForm);
                        if (e.StopChain)
                            break;
                    }
                }

                CurrentWord = "";
            }
            else if (e.KeyPressed.IsAlpha())
            {
                string newChar = e.KeyPressed switch
                {
                    Key.Oem3 => "ù",
                    Key.D2 => "é",
                    Key.D7 => "è",
                    Key.D9 => "ç",
                    Key.D0 => "à",
                    _ => e.KeyPressed.ToString().ToLower()
                };

                newChar = KeyboardStates.IsShifted() ? newChar.ToUpper() : newChar.ToLower();

                CurrentWord += newChar;

                Log.Verbose("New Char Pressed: {pressedChar}, Current Word: {@currentWord}", newChar, CurrentWord);
            }
            else if (e.KeyPressed == Key.Back && CurrentWord is not null && CurrentWord.Length > 0)
            {
                CurrentWord = CurrentWord.Remove(CurrentWord.Length - 1);
                Log.Verbose("Last char removed, Current Word: {@currentWord}", CurrentWord);
            }
            else
            {
                CurrentWord = "";
                Log.Verbose("Current Word Reset, Current Word: {@currentWord}", CurrentWord);
            }
        }

        void Choosing(object sender, KeyPressedArgs e)
        {
            if (e.KeyPressed is Key.Escape or Key.Space)
            {
                _currentListenerState = ListenerStates.Inserting;
                CurrentWord = "";
            }
        }

        public void Load(Action<string> currentWordCallback)
        {
            _listener.HookKeyboard(); 
            Continue();
            CurrentWordCallback = currentWordCallback;
        }
        public void Close() => _listener.UnHookKeyboard();
        public void Pause() => _listener.OnKeyPressed -= ListenerEvent;
        public void Continue() => _listener.OnKeyPressed += ListenerEvent;

        enum ListenerStates
        {
            Inserting,
            Choosing
        }
    }
}
