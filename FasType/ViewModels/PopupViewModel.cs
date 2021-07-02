using FasType.Core.Models;
using FasType.Core.Models.Dictionary;
using FasType.Core.Services;
using FasType.LLKeyboardListener;
using FasType.Models;
using FasType.Utils;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace FasType.ViewModels
{
    public class PopupViewModel : ObservableObject
    {
        bool _dropdownOpen;
        string _shortForm/*, _selectedString*/;
        ObservableCollection<BaseDictionaryElement> _collection;
        BaseDictionaryElement[]? _elements;
        BaseDictionaryElement? _selectedElement;
        Visibility _comboBoxVisibility, _busyIndicatorVisibility;

        readonly IRepositoriesManager _repositories;

        public bool DropdownOpen { get => _dropdownOpen; set => SetProperty(ref _dropdownOpen, value); }
        public Visibility ComboBoxVisibility { get => _comboBoxVisibility; set => SetProperty(ref _comboBoxVisibility, value); }
        public Visibility BusyIndicatorVisibility { get => _busyIndicatorVisibility; set => SetProperty(ref _busyIndicatorVisibility, value); }
        public ObservableCollection<BaseDictionaryElement> Collection { get => _collection; set => SetProperty(ref _collection, value); }
        public string ShortForm { get => _shortForm; set => SetProperty(ref _shortForm, value); }
        public Command<Window> CreateCommand { get; }
        public BaseDictionaryElement? SelectedElement { get => _selectedElement; set => SetProperty(ref _selectedElement, value); }
        //public string SelectedString { get => _selectedString; set => SetProperty(ref _selectedString, value); }


        public PopupViewModel(IRepositoriesManager repositories)
        {
            _repositories = repositories;
            CreateCommand = new(Create, CanCreate);

            Collection = new();
            _ = _collection ?? throw new NullReferenceException();

            ShortForm = string.Empty;
            _ = _shortForm ?? throw new NullReferenceException();

            ComboBoxVisibility = Visibility.Collapsed;
            BusyIndicatorVisibility = Visibility.Visible;
            DropdownOpen = false;
        }

        public void SearchForWord(string currentWord)
        {
            ShortForm = currentWord;
            currentWord = currentWord.ToLower();
            var l = _repositories.Linguistics.Words(currentWord).Select(s => "^" + s + "$").ToList();

            _elements = l.SelectMany(s => _repositories.Dictionary.GetElementsFromRegex(s) ?? Array.Empty<BaseDictionaryElement>()).Distinct().OrderBy(bde => bde.FullForm).ToArray();

            Collection = new ObservableCollection<BaseDictionaryElement>(_elements)
            {
                DictionaryElements.OtherElement,
                DictionaryElements.NoneElement
            };
            SelectedElement = Collection[0];

            BusyIndicatorVisibility = Visibility.Collapsed;
            ComboBoxVisibility = Visibility.Visible;
            DropdownOpen = true;
            CommandManager.InvalidateRequerySuggested();
        }

        bool CanCreate() => SelectedElement != null;//!string.IsNullOrEmpty(SelectedString);
        void Create(Window? w)
        {
            _ = w ?? throw new NullReferenceException();
            if (SelectedElement == DictionaryElements.NoneElement)//(SelectedString == Properties.Resources.None)
            {
                var r = MessageBox.Show(Properties.DialogResources.AddDictionary, Properties.Resources.Dictionary, MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (r == MessageBoxResult.Yes)
                    _repositories.Dictionary.Add(new SimpleDictionaryElement(ShortForm ?? throw new NullReferenceException(), string.Empty, string.Empty, string.Empty));
                w.Close();
                return;
            }

            var window = App.Current.ServiceProvider.GetRequiredService<Windows.AbbreviationWindow>();
            var page = App.Current.ServiceProvider.GetRequiredService<Pages.SimpleAbbreviationPage>();

            if (SelectedElement == DictionaryElements.OtherElement)//(SelectedString == Properties.Resources.Other)
            {
                page.SetNewAbbreviation(ShortForm ?? throw new NullReferenceException(), "", Array.Empty<string>());
            }
            else
            {

                page.SetNewAbbreviation(ShortForm ?? throw new NullReferenceException(), SelectedElement?.FullForm ?? throw new NullReferenceException(), SelectedElement.Others);
                //page.SetNewAbbreviation(ShortForm, SelectedString, Dictionary.GetElement(SelectedString)?.Others);
            }

            w.Hide();
            window.Content = page;
            if (window.ShowDialog() == true && window.DataContext is SimpleAbbreviationViewModel savm)
            {
                Core.Models.Abbreviations.SimpleAbbreviation abbrev = savm.CurrentAbbrev ?? throw new NullReferenceException();
                ((MainWindowViewModel)App.Current.MainWindow.DataContext).TryWriteAbbreviation(abbrev, ShortForm);
            }

            w.Close();
        }
    }
}
