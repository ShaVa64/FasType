using FasType.Models;
using FasType.Models.Abbreviations;
using FasType.Properties;
using FasType.Services;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace FasType.ViewModels
{
    public class SimpleAbbreviationViewModel : ObservableObject
    {
        readonly IDataStorage _storage;
        SimpleAbbreviation _currentAbbrev;
        string _shortForm, _fullForm, _genderForm, _pluralForm, _genderPluralForm;
        string _preview;

        public string ShortForm
        {
            get => _shortForm;
            set
            {
                if (SetProperty(ref _shortForm, value))
                    SetPreview();
            }
        }
        public string FullForm
        {
            get => _fullForm;
            set
            {
                if (SetProperty(ref _fullForm, value))
                    SetPreview();
            }
        }
        public string GenderForm
        {
            get => _genderForm;
            set
            {
                if (SetProperty(ref _genderForm, value))
                    SetPreview();
            }
        }
        public string PluralForm
        {
            get => _pluralForm;
            set
            {
                if (SetProperty(ref _pluralForm, value))
                    SetPreview();
            }
        }
        public string GenderPluralForm
        {
            get => _genderPluralForm;
            set
            {
                if (SetProperty(ref _genderPluralForm, value))
                    SetPreview();
            }
        }
        public string Preview { get => _preview; set => SetProperty(ref _preview, value); }

        public Command<Page> CreateNewCommand { get; set; }

        public SimpleAbbreviationViewModel(IDataStorage storage)
        {
            _storage = storage;
            _currentAbbrev = null;

            CreateNewCommand = new(CreateNew, CanCreateNew);

            ShortForm = "pss";
            FullForm = "passé";
            GenderForm = "passée";
            PluralForm = "passés";
            GenderPluralForm = "passées";
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
                var message = string.Format(Resources.AlreadyExistsErrorFormat, Environment.NewLine, _currentAbbrev.ElementaryRepresentation);
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
            CommandManager.InvalidateRequerySuggested();
            if (string.IsNullOrEmpty(ShortForm) || string.IsNullOrEmpty(FullForm))
                return;

            _currentAbbrev = new SimpleAbbreviation(ShortForm, FullForm, GenderForm, PluralForm, GenderPluralForm);

            Preview = _currentAbbrev.ComplexRepresentation;
        }
    }
}
