using FasType.Models;
using FasType.Utils;
using FasType.Core.Models.Abbreviations;
using FasType.Properties;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Extensions.DependencyInjection;
using FasType.Windows;
using FasType.Core.Models;
using FasType.Core.Services;
using FasType.Core.Models.Dictionary;

namespace FasType.ViewModels
{
    public class AddSimpleAbbreviationViewModel : SimpleAbbreviationViewModel
    {
        public AddSimpleAbbreviationViewModel(IRepositoriesManager repositories) : base(repositories, Resources.AddSimpleAbbrevTitle, Resources.Add, false)
        {
#if DEBUG
            //ShortForm = "pçé";
            //FullForm = "passé";
            //GenderForm = "passée";
            //PluralForm = "passés";
            //GenderPluralForm = "passées";
#endif
        }
        public AddSimpleAbbreviationViewModel(IRepositoriesManager repositories, string shortForm, string fullForm, string genderForm, string pluralForm, string genderPluralForm) : this(repositories)
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
            if (_repositories.Abbreviations.Contains(CurrentAbbrev))
            {
                var message = string.Format(DialogResources.AlreadyExistsErrorFormat, FullForm, Environment.NewLine);
                var res = MessageBox.Show(message, Resources.Error, MessageBoxButton.YesNo, MessageBoxImage.Error, MessageBoxResult.No);
                if (res == MessageBoxResult.No)
                    return;
                _repositories.Abbreviations.Update(CurrentAbbrev);
            }
            else
            {
                _repositories.Abbreviations.Add(CurrentAbbrev);
            }
            _repositories.Abbreviations.SaveChanges();

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
            
            if (_repositories.Abbreviations.Contains(CurrentAbbrev))
            {
                BorderBrush = Controls.BorderBrushTextBox.WarningBrush;
                FFToolTip = Resources.DupAbbrevToolTip;
                //return;
            }

            //TODO: Here
            //Preview = CurrentAbbrev.GetComplexRepresentation();
            //Preview = CurrentAbbrev.ComplexRepresentation;
        }
    }

    public class ModifySimpleAbbreviationViewModel : SimpleAbbreviationViewModel
    {
        readonly SimpleAbbreviation _toModify;

        public ModifySimpleAbbreviationViewModel(IRepositoriesManager repositories, SimpleAbbreviation sa) : base(repositories, Resources.ModifySimpleAbbrevTitle, Resources.Modify, true)
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
            _repositories.Abbreviations.Update(CurrentAbbrev);
            _repositories.Abbreviations.SaveChanges();

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
            //TODO: Here
            //Preview = CurrentAbbrev.ComplexRepresentation;
        }
    }

    public abstract class SimpleAbbreviationViewModel : ObservableObject
    {
        protected readonly IRepositoriesManager _repositories;

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

        public SimpleAbbreviationViewModel(IRepositoriesManager repositories, string title, string buttonText, bool shortFormReadOnly)
        {
            _repositories = repositories;
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

            PropertyChanged += SimpleAbbreviationViewModel_PropertyChanged;
        }

        void SimpleAbbreviationViewModel_PropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName?.EndsWith("Form", StringComparison.InvariantCulture) == true)
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
            if (_repositories.Dictionary.Contains(FullForm ?? throw new NullReferenceException()))
                return;

            var res = MessageBox.Show(DialogResources.AddDictionary, Resources.Dictionary, MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes);
            if (res == MessageBoxResult.No)
                return;

            //TODO: Add dictionary Add
            //_repositories.Dictionary.Add(CurrentAbbrev);            
        }

        void ComputeAutoComplete()
        {
            if (string.IsNullOrEmpty(FullForm))
            {
                GenderForm = PluralForm = GenderPluralForm = string.Empty;
                NotSkipping = true;
                return;
            }
            var elem = _repositories.Dictionary.GetById<SimpleDictionaryElement>(FullForm);
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
        }

        protected abstract void SetPreview();
    }
}
