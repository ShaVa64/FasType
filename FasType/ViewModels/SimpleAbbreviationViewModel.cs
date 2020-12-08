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
    public class SimpleAbbreviationViewModel : ObservableObject
    {
        static ILinguisticsStorage Linguistics => App.Current.ServiceProvider.GetRequiredService<ILinguisticsStorage>();
        static IAbbreviationStorage Storage => App.Current.ServiceProvider.GetRequiredService<IAbbreviationStorage>();

        SimpleAbbreviation _currentAbbrev;
        string _shortForm, _fullForm, _genderForm, _pluralForm, _genderPluralForm;
        string _sfToolTip, _ffToolTip, _preview;
        Brush _borderBrush;
        //bool _autoComplete;

        //public bool AutoComplete { get => _autoComplete; set => SetProperty(ref _autoComplete, value); }

        public string ShortForm { get => _shortForm; set => SetProperty(ref _shortForm, value/*, SetPreview*/); }
        public string FullForm { get => _fullForm; set => SetProperty(ref _fullForm, value/*, SetPreview*/); }
        public string GenderForm { get => _genderForm; set => SetProperty(ref _genderForm, value/*, SetPreview*/); }
        public string PluralForm { get => _pluralForm; set => SetProperty(ref _pluralForm, value/*, SetPreview*/); }
        public string GenderPluralForm { get => _genderPluralForm; set => SetProperty(ref _genderPluralForm, value/*, SetPreview*/); }
        
        public string SFToolTip { get => _sfToolTip; set => SetProperty(ref _sfToolTip, value); }
        public string FFToolTip { get => _ffToolTip; set => SetProperty(ref _ffToolTip, value); }
        
        public Brush BorderBrush { get => _borderBrush; set => SetProperty(ref _borderBrush, value); }

        public string Preview { get => _preview; set => SetProperty(ref _preview, value); }

        public Command<Page> CreateNewCommand { get; set; }
        public Command OpenLinguisticsCommand { get; set; }
        public Command AutoCompleteCommand { get; set; }

        public SimpleAbbreviationViewModel()
        {
            _currentAbbrev = null;

            CreateNewCommand = new(CreateNew, CanCreateNew);
            OpenLinguisticsCommand = new(OpenLinguistics, CanOpenLinguistics);
            AutoCompleteCommand = new(AutoComplete);

            ShortForm = FullForm = GenderForm = PluralForm = GenderPluralForm = string.Empty;
            SFToolTip = FFToolTip = null;
            BorderBrush = null;
            //AutoComplete = false;

            this.PropertyChanged += SimpleAbbreviationViewModel_PropertyChanged;
#if DEBUG
            //AutoComplete = true;
            ShortForm = "pss";
            //FullForm = "passé";
            //GenderForm = "passée";
            //PluralForm = "passés";
            //GenderPluralForm = "passées";
#endif
        }

        void SimpleAbbreviationViewModel_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (e.PropertyName.EndsWith("Form"))
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
        public void CreateNew(Page p)
        {
            if (_currentAbbrev == null || string.IsNullOrEmpty(_currentAbbrev.ShortForm) || string.IsNullOrEmpty(_currentAbbrev.FullForm))
            {
                MessageBox.Show(Resources.EmptyAbbrevDialog, Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                return;
            }

            if (Storage.Contains(_currentAbbrev))
            {
                var message = string.Format(Resources.AlreadyExistsErrorFormat, Environment.NewLine, FullForm);
                MessageBox.Show(message, Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                return;
            }
            if (!Storage.Add(_currentAbbrev))
            {
                var message = string.Format(Resources.ErrorDialogFormat, Environment.NewLine, _currentAbbrev.ElementaryRepresentation);
                MessageBox.Show(message, Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                return;
            }

            (p.Parent as Window).Close();
        }

        void ComputeAutoComplete()
        {
            if (string.IsNullOrEmpty(FullForm))
            {
                GenderForm = PluralForm = GenderPluralForm = string.Empty;
                return;
            }
            GenderForm = Linguistics.GenderCompletion.Grammarify(FullForm);// + "e";
            PluralForm = Linguistics.PluralCompletion.Grammarify(FullForm);// + "s";
            GenderPluralForm = Linguistics.GenderPluralCompletion.Grammarify(FullForm);// + "es";
        }

        void SetPreview()
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

            _currentAbbrev = new SimpleAbbreviation(ShortForm, FullForm, 0, GenderForm, PluralForm, GenderPluralForm);
            if (Storage.Contains(_currentAbbrev))
            {
                BorderBrush = Controls.BorderBrushTextBox.WarningBrush;
                FFToolTip = Resources.DupAbbrevToolTip;
                return;
            }

            Preview = _currentAbbrev.ComplexRepresentation;
        }
    }
}
