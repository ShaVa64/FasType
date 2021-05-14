using FasType.Models;
using FasType.Utils;
using FasType.Models.Abbreviations;
using FasType.Properties;
using FasType.Services;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Extensions.DependencyInjection;
using FasType.Windows;

namespace FasType.ViewModels
{
    public class AddSimpleAbbreviationViewModel : SimpleAbbreviationViewModel
    {
        public AddSimpleAbbreviationViewModel() : base(Resources.AddSimpleAbbrevTitle, Resources.Add, false)
        {
#if DEBUG
            //ShortForm = "pçé";
            //FullForm = "passé";
            //GenderForm = "passée";
            //PluralForm = "passés";
            //GenderPluralForm = "passées";
#endif
        }
        public AddSimpleAbbreviationViewModel(string shortForm, string fullForm, string genderForm, string pluralForm, string genderPluralForm) : this()
        {
            ShortForm = shortForm;
            FullForm = fullForm;
            GenderForm = genderForm;
            PluralForm = pluralForm;
            GenderPluralForm = genderPluralForm;
        }

        protected override void CreateNew(Page? p)
        {
            _ = p ?? throw new NullReferenceException();
            Window w = p.Parent as Window ?? throw new NullReferenceException();
            if (CurrentAbbrev == null || string.IsNullOrEmpty(CurrentAbbrev.ShortForm) || string.IsNullOrEmpty(CurrentAbbrev.FullForm))
            {
                MessageBox.Show(DialogResources.EmptyAbbrevDialog, Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                return;
            }
            if (Storage.Contains(CurrentAbbrev))
            {
                var message = string.Format(DialogResources.AlreadyExistsErrorFormat, FullForm, Environment.NewLine);
                var res = MessageBox.Show(message, Resources.Error, MessageBoxButton.YesNo, MessageBoxImage.Error, MessageBoxResult.No);
                if (res == MessageBoxResult.No)
                    return;
                var success = Storage.UpdateAbbreviation(CurrentAbbrev);
                if (!success)
                {
                    message = string.Format(DialogResources.ErrorDialogFormat, Environment.NewLine, CurrentAbbrev.ElementaryRepresentation);
                    MessageBox.Show(message, Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                    return;
                }
            }
            else if (!Storage.Add(CurrentAbbrev))
            {
                var message = string.Format(DialogResources.ErrorDialogFormat, Environment.NewLine, CurrentAbbrev.ElementaryRepresentation);
                MessageBox.Show(message, Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                return;
            }

            CheckDictionaryAdd();
            try
            {
                w.DialogResult = true;
            }
            catch { }
            finally
            {
                w.Close();
            }
        }

        protected override void SetPreview()
        {
            Preview = "";
            SFToolTip = FFToolTip = null;
            BorderBrush = null;
            CommandManager.InvalidateRequerySuggested();
            bool isSFEmpty = string.IsNullOrEmpty(ShortForm);
            bool isFFEmpty = string.IsNullOrEmpty(FullForm);
            if (isSFEmpty || isFFEmpty) 
            {
                if (isSFEmpty)
                    SFToolTip = Resources.EmptyAbbrevToolTip;
                if (isFFEmpty)
                    FFToolTip = Resources.EmptyAbbrevToolTip;
                return;
            }

            CurrentAbbrev = new SimpleAbbreviation(ShortForm ?? throw new NullReferenceException(),
                                                    FullForm ?? throw new NullReferenceException(),
                                                    0,
                                                    GenderForm ?? throw new NullReferenceException(),
                                                    PluralForm ?? throw new NullReferenceException(),
                                                    GenderPluralForm ?? throw new NullReferenceException());
            if (Storage.Contains(CurrentAbbrev))
            {
                BorderBrush = Controls.BorderBrushTextBox.WarningBrush;
                FFToolTip = Resources.DupAbbrevToolTip;
                //return;
            }

            Preview = CurrentAbbrev.ComplexRepresentation;
        }
    }

    public class ModifySimpleAbbreviationViewModel : SimpleAbbreviationViewModel
    {
        readonly SimpleAbbreviation _toModify;

        public ModifySimpleAbbreviationViewModel(SimpleAbbreviation sa) : base(Resources.ModifySimpleAbbrevTitle, Resources.Modify, true)
        {
            _toModify = sa;

            ShortForm = sa.ShortForm;
            FullForm = sa.FullForm;
            GenderForm = sa.GenderForm;
            PluralForm = sa.PluralForm;
            GenderPluralForm = sa.GenderPluralForm;
        }

        protected override void CreateNew(Page? p)
        {
            _ = p ?? throw new NullReferenceException();
            Window w = p.Parent as Window ?? throw new NullReferenceException();
            if (CurrentAbbrev == null || string.IsNullOrEmpty(CurrentAbbrev.ShortForm) || string.IsNullOrEmpty(CurrentAbbrev.FullForm))
            {
                MessageBox.Show(DialogResources.EmptyAbbrevDialog, Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                return;
            }
            if (!Storage.Remove(_toModify))
            {
                //var message = string.Format(Resources.AlreadyExistsErrorFormat, Environment.NewLine, FullForm);
                //var res = MessageBox.Show(message, Resources.Error, MessageBoxButton.YesNo, MessageBoxImage.Error, MessageBoxResult.No);
                //if (res == MessageBoxResult.No)
                //    return;
                //var success = Storage.UpdateAbbreviation(_currentAbbrev);
                //if (!success)
                //{
                //    message = string.Format(Resources.ErrorDialogFormat, Environment.NewLine, _currentAbbrev.ElementaryRepresentation);
                //    MessageBox.Show(message, Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                //    return;
                //}
                return;
            }
            if (!Storage.Add(CurrentAbbrev))
            {
                var message = string.Format(DialogResources.ErrorDialogFormat, Environment.NewLine, CurrentAbbrev.ElementaryRepresentation);
                MessageBox.Show(message, Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                return;
            }
            
            CheckDictionaryAdd();
            w.DialogResult = true;
            w.Close();
        }

        protected override void SetPreview()
        {
            Preview = "";
            SFToolTip = FFToolTip = null;
            BorderBrush = null;
            CommandManager.InvalidateRequerySuggested();
            bool isSFEmpty = string.IsNullOrEmpty(ShortForm);
            bool isFFEmpty = string.IsNullOrEmpty(FullForm);
            if (isSFEmpty || isFFEmpty)
            {
                if (isSFEmpty)
                    SFToolTip = Resources.EmptyAbbrevToolTip;
                if (isFFEmpty)
                    FFToolTip = Resources.EmptyAbbrevToolTip;
                return;
            }

            CurrentAbbrev = new SimpleAbbreviation(ShortForm ?? throw new NullReferenceException(),
                                                    FullForm ?? throw new NullReferenceException(),
                                                    0,
                                                    GenderForm ?? throw new NullReferenceException(),
                                                    PluralForm ?? throw new NullReferenceException(),
                                                    GenderPluralForm ?? throw new NullReferenceException());

            Preview = CurrentAbbrev.ComplexRepresentation;
        }
    }

    public abstract class SimpleAbbreviationViewModel : ObservableObject
    {
        protected static IDictionaryStorage Dictionary => App.Current.ServiceProvider.GetRequiredService<IDictionaryStorage>();
        protected static ILinguisticsStorage Linguistics => App.Current.ServiceProvider.GetRequiredService<ILinguisticsStorage>();
        protected static IAbbreviationStorage Storage => App.Current.ServiceProvider.GetRequiredService<IAbbreviationStorage>();

        public SimpleAbbreviation? CurrentAbbrev { get; protected set; }
        string? _shortForm, _fullForm, _genderForm, _pluralForm, _genderPluralForm;
        string? _sfToolTip, _ffToolTip, _preview;
        Brush? _borderBrush;
        bool _notSkipping;

        public string Title { get; }
        public string ButtonText { get; }
        public bool ShortFormReadOnly { get; }
        public bool NotSkipping { get => _notSkipping; set => SetProperty(ref _notSkipping, value); }
        public string? ShortForm { get => _shortForm; set => SetProperty(ref _shortForm, value/*, SetPreview*/); }
        public string? FullForm { get => _fullForm; set => SetProperty(ref _fullForm, value/*, SetPreview*/); }
        public string? GenderForm { get => _genderForm; set => SetProperty(ref _genderForm, value/*, SetPreview*/); }
        public string? PluralForm { get => _pluralForm; set => SetProperty(ref _pluralForm, value/*, SetPreview*/); }
        public string? GenderPluralForm { get => _genderPluralForm; set => SetProperty(ref _genderPluralForm, value/*, SetPreview*/); }
        
        public string? SFToolTip { get => _sfToolTip; set => SetProperty(ref _sfToolTip, value); }
        public string? FFToolTip { get => _ffToolTip; set => SetProperty(ref _ffToolTip, value); }
        
        public Brush? BorderBrush { get => _borderBrush; set => SetProperty(ref _borderBrush, value); }

        public string? Preview { get => _preview; set => SetProperty(ref _preview, value); }

        public Command<Page> CreateNewCommand { get; set; }
        public Command OpenLinguisticsCommand { get; set; }
        public Command AutoCompleteCommand { get; set; }

        public SimpleAbbreviationViewModel(string title, string buttonText, bool shortFormReadOnly)
        {
            Title = title;
            ShortFormReadOnly = shortFormReadOnly;
            ButtonText = buttonText;
            
            CurrentAbbrev = null;

            CreateNewCommand = new(CreateNew, CanCreateNew);
            OpenLinguisticsCommand = new(OpenLinguistics, CanOpenLinguistics);
            AutoCompleteCommand = new(AutoComplete);

            ShortForm = FullForm = GenderForm = PluralForm = GenderPluralForm = string.Empty;
            SFToolTip = FFToolTip = null;
            BorderBrush = null;
            NotSkipping = true;
            //AutoComplete = false;

            this.PropertyChanged += SimpleAbbreviationViewModel_PropertyChanged;
        }

        void SimpleAbbreviationViewModel_PropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName?.EndsWith("Form") == true)
                SetPreview();
            if (Settings.Default.FormsAutoComplete && e.PropertyName == nameof(FullForm))
                ComputeAutoComplete();
        }

        void AutoComplete()
        {
            if (Settings.Default.FormsAutoComplete)
                ComputeAutoComplete();
            else
                GenderForm = PluralForm = GenderPluralForm = string.Empty;
        }

        bool CanOpenLinguistics() => !LinguisticsWindow.IsOpen;
        void OpenLinguistics()
        {
            var lw = App.Current.ServiceProvider.GetRequiredService<LinguisticsWindow>();

            lw.Show();
        }

        bool CanCreateNew() => !string.IsNullOrEmpty(FullForm) && !string.IsNullOrEmpty(ShortForm);
        protected abstract void CreateNew(Page? p);
        protected void CheckDictionaryAdd()
        {
            if (Dictionary.Contains(FullForm ?? throw new NullReferenceException()))
                return;

            var res = MessageBox.Show(DialogResources.AddDictionary, Resources.Dictionary, MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes);
            if (res == MessageBoxResult.No)
                return;

            Dictionary.Add(CurrentAbbrev);
        }

        void ComputeAutoComplete()
        {
            if (string.IsNullOrEmpty(FullForm))
            {
                GenderForm = PluralForm = GenderPluralForm = string.Empty;
                NotSkipping = true;
                return;
            }
            //TODO: Get from dictionary
            var elem = Dictionary.GetElement<Models.Dictionary.SimpleDictionaryElement>(FullForm);
            if (elem == null)
            {
                GenderForm = PluralForm = GenderPluralForm = string.Empty;
                NotSkipping = true;
                return;
            }

            GenderForm = elem.GenderForm;
            PluralForm = elem.PluralForm;
            GenderPluralForm = elem.GenderPluralForm;
            NotSkipping = false;

            //GenderForm = Linguistics.GenderCompletion.Grammarify(FullForm);// + "e";
            //PluralForm = Linguistics.PluralCompletion.Grammarify(FullForm);// + "s";
            //GenderPluralForm = Linguistics.GenderPluralCompletion.Grammarify(FullForm);// + "es";
        }

        protected abstract void SetPreview();
    }
}
