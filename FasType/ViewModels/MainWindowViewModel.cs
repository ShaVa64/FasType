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
using FasType.Models.Abbreviations;
using FasType.Models;

namespace FasType.ViewModels
{
    public class MainWindowViewModel : ObservableObject, IKeyboardListenerHandler
    {
        string _currentWord;
        readonly LowLevelKeyboardListener _listener;
        ListenerStates _currentListenerState;
        readonly IAbbreviationStorage _storage;
        BaseAbbreviation _choosedAbbrev;
        List<BaseAbbreviation> _matchingAbbrevs;
        int _abbrevIndex;

        //string _choosedFullForm;
        //List<string> _matchingFullForms;
        //int _fullFormIndex;

        //public string ChoosedFullForm { get => _choosedFullForm; set => SetProperty(ref _choosedFullForm, value); }
        //public List<string> MatchingFullForms { get => _matchingFullForms; set => SetProperty(ref _matchingFullForms, value); }
        //public int FullFormIndex { get => _fullFormIndex; set => SetProperty(ref _fullFormIndex, value); }
        public static bool IsPaused => SeeAllWindow.IsOpen
                                       || AddAbbreviationWindow.IsOpen
                                       || LinguisticsWindow.IsOpen
                                       || AbbreviationMethodsWindow.IsOpen;

        public int AbbrevIndex { get => _abbrevIndex; set => SetProperty(ref _abbrevIndex, value); }
        public BaseAbbreviation ChoosedAbbrev { get => _choosedAbbrev; set => SetProperty(ref _choosedAbbrev, value); }
        public List<BaseAbbreviation> MatchingAbbrevs { get => _matchingAbbrevs; set => SetProperty(ref _matchingAbbrevs, value); }
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
        public Command<BaseAbbreviation> ChooseCommand { get; }
        public Command OpenLinguisticsCommand { get; }

        //static MainWindowViewModel() => _instance = App.Current.ServiceProvider.GetRequiredService<MainWindowViewModel>();
        public MainWindowViewModel(IAbbreviationStorage storage)
        {
            CurrentWord = string.Empty;
            _listener = new();
            _storage = storage;
            CurrentListenerState = ListenerStates.Inserting;

            AddNewCommand = new(AddNew, CanAddNew);
            SeeAllCommand = new(SeeAll, CanSeeAll);
            ChooseCommand = new(Choose, CanChoose);
            OpenLinguisticsCommand = new(OpenLinguistics, CanOpenLinguistics);
        }

        bool CanOpenLinguistics() => !LinguisticsWindow.IsOpen;
        void OpenLinguistics()
        {
            var lw = App.Current.ServiceProvider.GetRequiredService<LinguisticsWindow>();

            lw.Show();
        }

        bool CanAddNew(Type t) => t != null && t.IsSubclassOf(typeof(Page)) && !AddAbbreviationWindow.IsOpen;
        void AddNew(Type t)
        {
            var aaw = App.Current.ServiceProvider.GetRequiredService<AddAbbreviationWindow>();
            var p = App.Current.ServiceProvider.GetRequiredService(t) as Page;// Activator.CreateInstance(t) as Page;//new Pages.SimpleAbbreviationPage();

            aaw.Content = p;
            aaw.Show();
        }

        bool CanSeeAll() => _storage.Count > 0 && !SeeAllWindow.IsOpen;
        void SeeAll()
        {
            var saw = App.Current.ServiceProvider.GetRequiredService<SeeAllWindow>();

            saw.Show();
        }

        #region IKeyboardListenerHandler
        bool CanChoose() => CurrentListenerState == ListenerStates.Choosing;
        void Choose(BaseAbbreviation abbrev)
        {
            TryWriteAbbreviation(abbrev, CurrentWord, plusOne: true);

            //ChoosedFullForm = null;
            //MatchingFullForms = null;
            ChoosedAbbrev = null;
            MatchingAbbrevs = null;
            CurrentListenerState = ListenerStates.Inserting;
            CurrentWord = "";
        }

        bool TryWriteAbbreviation(BaseAbbreviation abbrev, string shortForm, bool plusOne = false)
        {
            if (abbrev.TryGetFullForm(shortForm, out string fullForm))
            {
                string word = CurrentWord.IsFirstCharUpper() ? fullForm.FirstCharToUpper() : fullForm;
                Input.Erase(CurrentWord.Length + (plusOne ? 1 : 0));
                Input.TextEntry(word + " ");
                _storage.UpdateUsed(abbrev);
                return true;
            }
            return false;
        }

        void ListenerEvent(object sender, KeyPressedEventArgs e)
        {
            if (IsPaused) 
                return;
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

                //MatchingFullForms = abbrevs.Select(a => a.GetFullForm(shortForm)).ToList();
                //ChoosedFullForm = MatchingFullForms[0];
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
                string newChar = (e.Old?.KeyPressed, e.Old?.IsShifted, e.KeyPressed, e.IsShifted) switch
                {
                    (Key.Oem6, false, Key.E   , false) => "ê",
                    (Key.Oem6, true , Key.E   , false) => "ë",
                    (Key.Oem6, false, Key.E   , true ) => "Ê",
                    (Key.Oem6, true , Key.E   , true ) => "Ë",
                    (Key.Oem6, false, Key.U   , false) => "û",
                    (Key.Oem6, true , Key.U   , false) => "ü",
                    (Key.Oem6, false, Key.U   , true ) => "Û",
                    (Key.Oem6, true , Key.U   , true ) => "Ü",
                    (_       , _    , Key.Oem3, false) => "ù",
                    (_       , _    , Key.D2  , false) => "é",
                    (_       , _    , Key.D7  , false) => "è",
                    (_       , _    , Key.D9  , false) => "ç",
                    (_       , _    , Key.D0  , false) => "à",
                    (_       , _    , _       , false) => e.KeyPressed.ToString().ToLower(),
                    (_       , _    , _       , true ) => e.KeyPressed.ToString(),
                };

                //newChar = e.IsShifted ? newChar.ToUpper() : newChar.ToLower();
                CurrentWord += newChar.Single();

                Log.Verbose("New Char Pressed: {pressedChar}, Current Word: {@currentWord}", newChar, CurrentWord);
            }
            else if (e.KeyPressed == Key.Back && !string.IsNullOrEmpty(CurrentWord))
            {
                CurrentWord = CurrentWord.Remove(CurrentWord.Length - 1);
                Log.Verbose("Last char removed, Current Word: {@currentWord}", CurrentWord);
            }
            else if (e.KeyPressed.IsModifier() || (e.KeyPressed == Key.Oem6 && e.Old.KeyPressed != Key.Oem6)) { }
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

                //string word = CurrentWord.IsFirstCharUpper() ? ChoosedFullForm.FirstCharToUpper() : ChoosedFullForm;
                //Input.Erase(CurrentWord.Length + 1);
                //Input.TextEntry(word + " ");
                //ChoosedFullForm = null;
                //MatchingFullForms = null;
                //CurrentWord = "";
                bool b = TryWriteAbbreviation(ChoosedAbbrev, CurrentWord, plusOne: true);
                if (b)
                {
                    ChoosedAbbrev = null;
                    MatchingAbbrevs = null;
                    CurrentWord = "";
                }
                else
                {
                    CurrentListenerState = ListenerStates.Choosing;
                }
            }
            else if (e.KeyPressed is Key.Down)
            {
                //if (FullFormIndex < MatchingFullForms.Count - 1)
                //    FullFormIndex++;
                if (AbbrevIndex < MatchingAbbrevs.Count - 1)
                    AbbrevIndex++;
            }
            else if (e.KeyPressed is Key.Up)
            {
                //if (FullFormIndex > 0)
                //    FullFormIndex--;
                if (AbbrevIndex > 0)
                    AbbrevIndex--;
            }
            else //if (e.KeyPressed is Key.Escape or Key.Space)
            {
                //ChoosedFullForm = null;
                //MatchingFullForms = null;
                ChoosedAbbrev = null;
                MatchingAbbrevs = null;
                CurrentListenerState = ListenerStates.Inserting;
                CurrentWord = "";
            }
        }

        //public void Load() => Load(null, null);
        public void Load(object sender, RoutedEventArgs e)
        {
            _listener.HookKeyboard();
            _listener.OnKeyPressed += ListenerEvent;
        }
        //public void Close() => Load(null, null);
        public void Close(object sender, CancelEventArgs e)
        {
            _listener.OnKeyPressed -= ListenerEvent;
            _listener.UnHookKeyboard();
        }

        //public void Pause()
        //{
        //    _listener.OnKeyPressed -= ListenerEvent;
        //}

        //public void Continue()
        //{
        //    if (IsPaused)
        //    {
        //        _listener.OnKeyPressed += ListenerEvent;
        //    }
        //}

        enum ListenerStates
        {
            Inserting,
            Choosing
        }
        #endregion
    }
}
