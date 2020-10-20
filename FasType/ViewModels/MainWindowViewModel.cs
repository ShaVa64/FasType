using FasType.LLKeyboardListener;
using FasType.Windows;
using FasType.Services;
using FasType.Utils;
using Microsoft.Extensions.DependencyInjection;
using Serilog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WindowsInput;
using FasType.Models.Abbreviations;
using FasType.Models;

namespace FasType.ViewModels
{
    public class MainWindowViewModel : BaseViewModel, IKeyboardListenerHandler
    {
        string _currentWord;
        readonly LowLevelKeyboardListener _listener;
        readonly InputSimulator _sim;
        ListenerStates _currentListenerState;
        readonly IDataStorage _storage;
        IAbbreviation _choosedAbbrev;
        List<IAbbreviation> _matchingAbbrevs;
        int _abbrevIndex;
        
        public int AbbrevIndex
        {
            get => _abbrevIndex;
            set => SetProperty(ref _abbrevIndex, value);
        }
        public IAbbreviation ChoosedAbbrev
        {
            get => _choosedAbbrev;
            set => SetProperty(ref _choosedAbbrev, value);
        }
        public List<IAbbreviation> MatchingAbbrevs
        {
            get => _matchingAbbrevs;
            set => SetProperty(ref _matchingAbbrevs, value);
        }
        ListenerStates CurrentListenerState 
        {
            get => _currentListenerState; 
            set 
            {
                if (SetProperty(ref _currentListenerState, value))
                    OnPropertyChanged(nameof(IsChoosing));
            }
        }
        public bool IsChoosing => CurrentListenerState == ListenerStates.Choosing;
        public string CurrentWord { get => _currentWord; private set => SetProperty(ref _currentWord, value); }
        public Command<Type> AddNewCommand { get; }
        public Command SeeAllCommand { get; }

        public MainWindowViewModel(IDataStorage storage)
        {
            _listener = new LowLevelKeyboardListener();
            _sim = new InputSimulator();
            _storage = storage;
            CurrentListenerState = ListenerStates.Inserting;
            AddNewCommand = new(AddNew, CanAddNew);
            SeeAllCommand = new(SeeAll, CanSeeAll);
        }

        bool CanAddNew() => true;
        void AddNew(Type t)
        {
            var tw = App.Current.ServiceProvider.GetRequiredService<AddAbbreviationWindow>();
            var p = App.Current.ServiceProvider.GetRequiredService(t) as Page;// Activator.CreateInstance(t) as Page;//new Pages.SimpleAbbreviationPage();

            tw.Content = p;

            Pause();
            tw.ShowDialog();
            Continue();
        }

        bool CanSeeAll() => _storage.Count > 0;
        public void SeeAll()
        {
            var tw = App.Current.ServiceProvider.GetRequiredService<SeeAllWindow>();

            Pause();
            tw.ShowDialog();
            Continue();
        }

        #region IKeyboardListenerHandler
        bool TryWriteAbbreviation(IAbbreviation abbrev, string shortForm, bool plusOne = false)
        {
            if (abbrev.TryGetFullForm(shortForm, out string fullForm))
            {
                string word = CurrentWord.IsFirstCharUpper() ? fullForm.FirstCharToUpper() : fullForm;
                _sim.Keyboard.KeyPress(Enumerable.Repeat(WindowsInput.Native.VirtualKeyCode.BACK, CurrentWord.Length + (plusOne ? 1 : 0)).ToArray());
                _sim.Keyboard.TextEntry(word + " ");
                return true;
            }
            return false;
        }

        void ListenerEvent(object sender, KeyPressedEventArgs e)
        {
            Log.Information("Current Listener State: {listenerState}", CurrentListenerState);
            if (CurrentListenerState is ListenerStates.Inserting)
                Inserting(sender, e);
            else if (CurrentListenerState is ListenerStates.Choosing)
                Choosing(sender, e);
        }

        void Inserting(object sender, KeyPressedEventArgs e)
        {
            if (e.KeyPressed == Key.Space)
            {
                string shortForm = CurrentWord.ToLower();

                var abbrevs = _storage[shortForm].ToList();

                if (abbrevs.Count == 0)
                {
                    CurrentWord = "";
                    return;
                }

                if (abbrevs.Count == 1)
                {
                    var abbrev = abbrevs.Single();
                    e.StopChain |= TryWriteAbbreviation(abbrev, shortForm);
                    CurrentWord = "";
                    return;
                }
                //else if (abbrevs.Count > 1)
                CurrentListenerState = ListenerStates.Choosing;


                MatchingAbbrevs = abbrevs;
                ChoosedAbbrev = MatchingAbbrevs[0];
                //foreach (var abbrev in abbrevs)
                //{
                //    e.StopChain |= TryWriteAbbreviation(abbrev, shortForm);
                //    if (e.StopChain)
                //        break;
                //}
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

        void Choosing(object sender, KeyPressedEventArgs e)
        {
            e.StopChain = true;
            if (e.KeyPressed is Key.Enter)
            {
                CurrentListenerState = ListenerStates.Inserting;

                bool b = TryWriteAbbreviation(ChoosedAbbrev, CurrentWord, plusOne: true);
                if (b)
                {
                    ChoosedAbbrev = null;
                    MatchingAbbrevs = null;
                    CurrentWord = "";
                }
                else
                    CurrentListenerState = ListenerStates.Choosing;
            }
            else if (e.KeyPressed is Key.Down)
            {
                if (AbbrevIndex < MatchingAbbrevs.Count - 1)
                    AbbrevIndex++;
            }
            else if (e.KeyPressed is Key.Up)
            {
                if (AbbrevIndex > 0)
                    AbbrevIndex--;
            }
            else //if (e.KeyPressed is Key.Escape or Key.Space)
            {
                ChoosedAbbrev = null;
                MatchingAbbrevs = null;
                CurrentListenerState = ListenerStates.Inserting;
                CurrentWord = "";
            }
        }

        public void Load() => Load(null, null);
        public void Load(object sender, RoutedEventArgs e)
        {
            _listener.HookKeyboard();
            Continue();
        }
        public void Close() => Load(null, null);
        public void Close(object sender, CancelEventArgs e) => _listener.UnHookKeyboard();
        public void Pause() => _listener.OnKeyPressed -= ListenerEvent;
        public void Continue() => _listener.OnKeyPressed += ListenerEvent;

        enum ListenerStates
        {
            Inserting,
            Choosing
        }
        #endregion
    }
}
