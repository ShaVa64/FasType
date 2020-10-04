using FasType.Models;
using FasType.Services;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace FasType.ViewModels
{
    public class ToolWindowViewModel : BaseViewModel
    {
        readonly IDataStorage _storage;
        SimpleAbbreviation _currentAbbrev;
        string _shortForm, _fullForm;
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
        public string Preview { get => _preview; set => SetProperty(ref _preview, value); }

        public RoutedCommand CreateNewCommand { get; set; }

        public ToolWindowViewModel(IDataStorage storage)
        {
            _storage = storage;
            _currentAbbrev = null;

            CreateNewCommand = new("CreateNew", typeof(ToolWindowViewModel));            
        }

        public void CreateNew(object sender, ExecutedRoutedEventArgs e)
        {
            if (_currentAbbrev == null || string.IsNullOrEmpty(_currentAbbrev.ShortForm) || string.IsNullOrEmpty(_currentAbbrev.FullForm))
            {
                MessageBox.Show("You can't create an empty abbreviation", "Error", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                return;
            }

            bool b = _storage.Add(_currentAbbrev);
            if (!b)
            {
                MessageBox.Show($"An error has occured while trying to create the abbreviation ({_currentAbbrev.ElementaryRepresentation}).", "Error", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                return;
            }

            ((sender as Page).Parent as Window).Close();
        }

        public void CanCreateNew(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = !(string.IsNullOrEmpty(ShortForm) || string.IsNullOrEmpty(FullForm));
        }

        void SetPreview()
        {
            Preview = "";
            CommandManager.InvalidateRequerySuggested();
            if (string.IsNullOrEmpty(ShortForm) || string.IsNullOrEmpty(FullForm))
                return;

            _currentAbbrev = new SimpleAbbreviation(ShortForm, FullForm);

            Preview = _currentAbbrev.ComplexRepresentation;
        }
    }
}
