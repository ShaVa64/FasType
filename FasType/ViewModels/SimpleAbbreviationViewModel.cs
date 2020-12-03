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

namespace FasType.ViewModels
{
    public class SimpleAbbreviationViewModel : ObservableObject
    {
        readonly IAbbreviationStorage _storage;
        static readonly Brush _defaultBorderBrush = new SolidColorBrush(Color.FromRgb(170, 170, 170));
        static readonly Brush _duplicateBorderBrush = Brushes.DarkOrange;
        SimpleAbbreviation _currentAbbrev;
        string _shortForm, _fullForm, _genderForm, _pluralForm, _genderPluralForm;
        string _sfToolTip, _ffToolTip, _preview;
        Brush _borderBrush;

        public string ShortForm { get => _shortForm; set => SetProperty(ref _shortForm, value, SetPreview); }
        public string FullForm { get => _fullForm; set => SetProperty(ref _fullForm, value, SetPreview); }
        public string GenderForm { get => _genderForm; set => SetProperty(ref _genderForm, value, SetPreview); }
        public string PluralForm { get => _pluralForm; set => SetProperty(ref _pluralForm, value, SetPreview); }
        public string GenderPluralForm { get => _genderPluralForm; set => SetProperty(ref _genderPluralForm, value, SetPreview); }
        
        public string SFToolTip { get => _sfToolTip; set => SetProperty(ref _sfToolTip, value); }
        public string FFToolTip { get => _ffToolTip; set => SetProperty(ref _ffToolTip, value); }
        
        public Brush BorderBrush { get => _borderBrush; set => SetProperty(ref _borderBrush, value); }

        public string Preview { get => _preview; set => SetProperty(ref _preview, value); }

        public Command<Page> CreateNewCommand { get; set; }

        public SimpleAbbreviationViewModel(IAbbreviationStorage storage)
        {
            _storage = storage;
            _currentAbbrev = null;

            CreateNewCommand = new(CreateNew, CanCreateNew);

            ShortForm = FullForm = GenderForm = PluralForm = GenderPluralForm = string.Empty;
            SFToolTip = FFToolTip = null;
            BorderBrush = _defaultBorderBrush;
#if DEBUG
            ShortForm = "pss";
            FullForm = "passé";
            //GenderForm = "passée";
            //PluralForm = "passés";
            //GenderPluralForm = "passées";
#endif
        }

        bool CanCreateNew() => !string.IsNullOrEmpty(FullForm) && !string.IsNullOrEmpty(ShortForm);
        public void CreateNew(Page p)
        {
            if (_currentAbbrev == null || string.IsNullOrEmpty(_currentAbbrev.ShortForm) || string.IsNullOrEmpty(_currentAbbrev.FullForm))
            {
                MessageBox.Show(Resources.EmptyAbbrevDialog, Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                return;
            }

            if (_storage.Contains(_currentAbbrev))
            {
                var message = string.Format(Resources.AlreadyExistsErrorFormat, Environment.NewLine, FullForm);
                MessageBox.Show(message, Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                return;
            }
            if (!_storage.Add(_currentAbbrev))
            {
                var message = string.Format(Resources.ErrorDialogFormat, Environment.NewLine, _currentAbbrev.ElementaryRepresentation);
                MessageBox.Show(message, Resources.Error, MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                return;
            }

            (p.Parent as Window).Close();
        }

        void SetPreview()
        {
            Preview = "";
            SFToolTip = FFToolTip = null;
            BorderBrush = _defaultBorderBrush;
            CommandManager.InvalidateRequerySuggested();
            bool isSFEmpty = string.IsNullOrEmpty(ShortForm);
            bool isFFEmpty = string.IsNullOrEmpty(FullForm);
            if (isSFEmpty || isFFEmpty) 
            {
                if (isSFEmpty)
                    SFToolTip = "Cannot create an empty abbreviation.";
                if (isFFEmpty)
                    FFToolTip = "Cannot create an empty abbreviation.";
                return;
            }

            _currentAbbrev = new SimpleAbbreviation(ShortForm, FullForm, 0, GenderForm, PluralForm, GenderPluralForm);
            if (_storage.Contains(_currentAbbrev))
            {
                BorderBrush = _duplicateBorderBrush;
                FFToolTip = "Such abbreviation already exists.";
                return;
            }

            Preview = _currentAbbrev.ComplexRepresentation;
        }
    }
}
