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
using System.Threading.Channels;
using System.Threading;
using System.Threading.Tasks;

namespace FasType.ViewModels
{
    public class MainWindowViewModel : ObservableObject, IKeyboardListenerHandler
    {
        string _currentWord;
        readonly LowLevelKeyboardListener _listener;
        ListenerStates _currentListenerState;
        readonly IAbbreviationStorage _storage;
        readonly IDictionaryStorage _dictionary;
        BaseAbbreviation? _choosedAbbrev;
        List<BaseAbbreviation>? _matchingAbbrevs;
        int _abbrevIndex;
        //System.Windows.Media.Brush _background;

        //string _choosedFullForm;
        //List<string> _matchingFullForms;
        //int _fullFormIndex;

        //public string ChoosedFullForm { get => _choosedFullForm; set => SetProperty(ref _choosedFullForm, value); }
        //public List<string> MatchingFullForms { get => _matchingFullForms; set => SetProperty(ref _matchingFullForms, value); }
        //public int FullFormIndex { get => _fullFormIndex; set => SetProperty(ref _fullFormIndex, value); }

        public static bool IsPaused => SeeAllWindow.IsOpen
                                       || AbbreviationWindow.IsOpen
                                       || LinguisticsWindow.IsOpen
                                       || AbbreviationMethodsWindow.IsOpen
                                       || OneLettersWindow.IsOpen
                                       || PopupWindow.IsOpen;

        public int AbbrevIndex { get => _abbrevIndex; set => SetProperty(ref _abbrevIndex, value); }
        public BaseAbbreviation? ChoosedAbbrev { get => _choosedAbbrev; set => SetProperty(ref _choosedAbbrev, value); }
        public List<BaseAbbreviation>? MatchingAbbrevs { get => _matchingAbbrevs; set => SetProperty(ref _matchingAbbrevs, value); }
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
        public Command ChooseCommand { get; }
        public Command OpenLinguisticsCommand { get; }
        public Command<System.Media.SystemSound> PlaySoundCommand { get; }
        //public System.Windows.Media.Brush Background { get => _background; set => SetProperty(ref _background, value); }

        //static MainWindowViewModel() => _instance = App.Current.ServiceProvider.GetRequiredService<MainWindowViewModel>();
        public MainWindowViewModel(IAbbreviationStorage storage, IDictionaryStorage dictionary)
        {
            CurrentWord = string.Empty;
            _ = _currentWord ?? throw new NullReferenceException();
            _listener = new();
            _storage = storage;
            _dictionary = dictionary;
            CurrentListenerState = ListenerStates.Inserting;

            //Background = System.Windows.Media.Brushes.White;
            //_ = _background ?? throw new NullReferenceException();

            AddNewCommand = new(AddNew, CanAddNew);
            SeeAllCommand = new(SeeAll, CanSeeAll);
            ChooseCommand = new(Choose, CanChoose);
            OpenLinguisticsCommand = new(OpenLinguistics, CanOpenLinguistics);
            PlaySoundCommand = new(PlaySound);
        }

        static void PlaySound(System.Media.SystemSound? sound) => sound?.Play();

        bool CanOpenLinguistics() => !LinguisticsWindow.IsOpen;
        void OpenLinguistics()
        {
            var lw = App.Current.ServiceProvider.GetRequiredService<LinguisticsWindow>();

            lw.Show();
        }

        bool CanAddNew(Type? t) => t != null && t.IsSubclassOf(typeof(Page)) && !AbbreviationWindow.IsOpen;
        void AddNew(Type? t)
        {
            _ = t ?? throw new NullReferenceException();
            var aaw = App.Current.ServiceProvider.GetRequiredService<AbbreviationWindow>();
            var p = App.Current.ServiceProvider.GetRequiredService(t) as Page;

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
        void Choose()
        {
            TryWriteAbbreviation(ChoosedAbbrev ?? throw new NullReferenceException(), CurrentWord);

            CurrentListenerState = ListenerStates.Inserting;
            ChoosedAbbrev = null;
            MatchingAbbrevs = null;
            CurrentWord = "";
            App.Current.MainWindow.Hide();
        }

        bool TryWriteAbbreviation(BaseAbbreviation abbrev, string shortForm)
        {
            if (abbrev.TryGetFullForm(shortForm, out string? fullForm))
            {
                _ = fullForm ?? throw new NullReferenceException();
                string word = CurrentWord.IsFirstCharUpper() ? fullForm.FirstCharToUpper() : fullForm;
                _listener.OnKeyPressed -= ListenerEvent;
                Input.Erase(CurrentWord.Length);
                Input.TextEntry(word + ' ');
                _listener.OnKeyPressed += ListenerEvent;
                _storage.UpdateUsed(abbrev);
                return true;
            }
            return false;
        }

        void ListenerEvent(object? sender, KeyPressedEventArgs e)
        {
            if (IsPaused)
                return;
            Log.Information("Current Listener State: {listenerState}", CurrentListenerState);
            switch (CurrentListenerState)
            {
                case ListenerStates.Inserting:
                    Inserting(sender, e);
                    break;
                case ListenerStates.Choosing:
                    Choosing(sender, e);
                    break;
                default:
                    throw new NotImplementedException();
            }
        }
        void StartWindowAlert()
        {
            new System.Media.SoundPlayer(@"Assets\sound.wav").Play();
            while (App.Current.FlashApp() == false);

            //Background = System.Windows.Media.Brushes.Red;
            //await System.Threading.Tasks.Task.Delay(400);
            //Background = System.Windows.Media.Brushes.White;
        }

        static void StopWindowAlert()
        {
            while (App.Current.StopFlashingApp());
        }
        void Inserting(object? sender, KeyPressedEventArgs e)
        {
            if (e.KeyPressed == Key.Space && !string.IsNullOrEmpty(CurrentWord))
            {
                e.StopChain = true;
                string shortForm = CurrentWord.ToLower();

                var abbrevs = _storage[shortForm].ToList();

                if (abbrevs.Count == 0)
                {
                    //var vals = App.Current.ServiceProvider.GetRequiredService<ILinguisticsStorage>().Words(CurrentWord);
                    //var dict = App.Current.ServiceProvider.GetRequiredService<IDictionaryStorage>();

                    //var elems = vals.Select(val => dict.GetElement(val)).Where(elem => elem != null).ToList();

                    //if (elems.Count > 0)
                    //{

                    //}

                    //using var dict = App.Current.ServiceProvider.GetRequiredService<IDictionaryStorage>();
                    if (Properties.Settings.Default.AbbrevsAutoCreation && !_dictionary.Contains(shortForm))
                    {
                        //var window = App.Current.ServiceProvider.GetRequiredService<PopupWindow>();
                        //window.SearchForWord(shortForm);
                        //window.Show();
                    }

                    CurrentWord = "";
                    return;
                }

                if (abbrevs.Count == 1)
                {
                    var abbrev = abbrevs[0];
                    var couldWrite = TryWriteAbbreviation(abbrev, shortForm);
                    CurrentWord = "";
                    return;
                }
                //else if (abbrevs.Count > 1)
                CurrentListenerState = ListenerStates.Choosing;

                //MatchingFullForms = abbrevs.Select(a => a.GetFullForm(shortForm)).ToList();
                //ChoosedFullForm = MatchingFullForms[0];

                App.Current.MainWindow.Show();
                //var p = Caret.GetCaretPos();
                //App.Current.MainWnd.ShowAt(p);
                StartWindowAlert();
                MatchingAbbrevs = abbrevs.OrderByDescending(a => a.Used).Append(BaseAbbreviation.OtherAbbreviation).ToList();
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
                    (Key.Oem6, false, Key.O   , false) => "ô",
                    (Key.Oem6, true , Key.O   , false) => "ö",
                    (Key.Oem6, false, Key.O   , true ) => "Ô",
                    (Key.Oem6, true , Key.O   , true ) => "Ö",
                    (Key.Oem6, false, Key.A   , false) => "â",
                    (Key.Oem6, true , Key.A   , false) => "ä",
                    (Key.Oem6, false, Key.A   , true ) => "Â",
                    (Key.Oem6, true , Key.A   , true ) => "Ä",
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

                Log.Verbose("New Char Pressed: '{pressedChar}' ({vkCode}), Current Word: \"{@currentWord}\"", newChar, e.New.VkCode, CurrentWord);
            }
            else if (e.KeyPressed == Key.Back && !string.IsNullOrEmpty(CurrentWord))
            {
                CurrentWord = CurrentWord[..^1];//.Remove(CurrentWord.Length - 1);
                Log.Verbose("Last char removed, Current Word: \"{@currentWord}\"", CurrentWord);
            }
            else if (e.KeyPressed.IsModifier() || (e.KeyPressed == Key.Oem6 && e.Old?.KeyPressed != Key.Oem6)) { }
            else 
            {
                CurrentWord = "";
                Log.Verbose("Current Word Reset ({char} {vkCode}), Current Word: \"{@currentWord}\"", e.KeyPressed, e.New.VkCode, CurrentWord);
            }
        }

        void Choosing(object? sender, KeyPressedEventArgs e)
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
                if (ChoosedAbbrev == BaseAbbreviation.OtherAbbreviation)
                {
                    var aaw = App.Current.ServiceProvider.GetRequiredService<AbbreviationWindow>();
                    var p = App.Current.ServiceProvider.GetRequiredService<Pages.SimpleAbbreviationPage>();

                    aaw.Content = p;
                    p.SetNewAbbreviation(CurrentWord, "", Array.Empty<string>());
                    aaw.Show();
                    aaw.Activate();

                    StopWindowAlert();
                    ChoosedAbbrev = null;
                    MatchingAbbrevs = null;
                    CurrentWord = "";
                    App.Current.MainWindow.Hide();
                    return;
                }

                bool b = TryWriteAbbreviation(ChoosedAbbrev ?? throw new NullReferenceException(), CurrentWord);
                if (b)
                {
                    ChoosedAbbrev = null;
                    MatchingAbbrevs = null;
                    CurrentWord = "";
                    App.Current.MainWindow.Hide();
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
                if (AbbrevIndex < MatchingAbbrevs!.Count - 1)
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
                App.Current.MainWindow.Hide();
            }
        }

        //public void Load() => Load(null, null);
        public void Load()
        {
            _listener.HookKeyboard();
            _listener.OnKeyPressed += ListenerEvent;
        }
        //public void Close() => Load(null, null);
        public void Close()
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
