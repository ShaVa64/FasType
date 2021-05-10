using FasType.LLKeyboardListener;
using FasType.Models;
using FasType.Models.Dictionary;
using FasType.Services;
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
        string _shortForm/*, _selectedString*/;
        ObservableCollection<BaseDictionaryElement> _collection;
        BaseDictionaryElement[]? _elements;
        BaseDictionaryElement? _selectedElement;
        Visibility _comboBoxVisibility, _busyIndicatorVisibility;

        //static Services.ILinguisticsStorage Linguistics => App.Current.ServiceProvider.GetRequiredService<Services.ILinguisticsStorage>();
        //static Services.IDictionaryStorage Dictionary => App.Current.ServiceProvider.GetRequiredService<Services.IDictionaryStorage>();

        readonly ILinguisticsStorage _linguistics;
        readonly IDictionaryStorage _dictionary;

        public Visibility ComboBoxVisibility { get => _comboBoxVisibility; set => SetProperty(ref _comboBoxVisibility, value); }
        public Visibility BusyIndicatorVisibility { get => _busyIndicatorVisibility; set => SetProperty(ref _busyIndicatorVisibility, value); }
        public ObservableCollection<BaseDictionaryElement> Collection { get => _collection; set => SetProperty(ref _collection, value); }
        public string ShortForm { get => _shortForm; set => SetProperty(ref _shortForm, value); }
        public Command<Window> CreateCommand { get; }
        public BaseDictionaryElement? SelectedElement { get => _selectedElement; set => SetProperty(ref _selectedElement, value); }
        //public string SelectedString { get => _selectedString; set => SetProperty(ref _selectedString, value); }


        public PopupViewModel(ILinguisticsStorage linguistics, IDictionaryStorage dictionary)
        {
            _linguistics = linguistics;
            _dictionary = dictionary;
            CreateCommand = new(Create, CanCreate);

            Collection = new();
            _ = _collection ?? throw new NullReferenceException();

            ShortForm = string.Empty;
            _ = _shortForm ?? throw new NullReferenceException();

            ComboBoxVisibility = Visibility.Collapsed;
            BusyIndicatorVisibility = Visibility.Visible;
        }

        public void SearchForWord(string currentWord)
        {
            ShortForm = currentWord;
            currentWord = currentWord.ToLower();
            var l = _linguistics.Words(currentWord).Select(s => "^" + s + "$").ToList();

            _elements = l.SelectMany(s => _dictionary.GetElements(s) ?? Array.Empty<BaseDictionaryElement>()).Distinct().ToArray();

            Collection = new ObservableCollection<BaseDictionaryElement>(_elements)
            {
                BaseDictionaryElement.OtherElement,
                BaseDictionaryElement.NoneElement
            };
            SelectedElement = Collection[0];

            BusyIndicatorVisibility = Visibility.Collapsed;
            ComboBoxVisibility = Visibility.Visible;
            CommandManager.InvalidateRequerySuggested();
        }

        bool CanCreate() => SelectedElement != null;//!string.IsNullOrEmpty(SelectedString);
        void Create(Window? w)
        {
            _ = w ?? throw new NullReferenceException();
            if (SelectedElement == BaseDictionaryElement.NoneElement)//(SelectedString == Properties.Resources.None)
            {
                var r = MessageBox.Show(Properties.DialogResources.AddDictionary, Properties.Resources.Dictionary, MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (r == MessageBoxResult.Yes)
                    _dictionary.Add(new SimpleDictionaryElement(ShortForm ?? throw new NullReferenceException(), string.Empty, string.Empty, string.Empty));
                w.Close();
                return;
            }

            var window = App.Current.ServiceProvider.GetRequiredService<Windows.AbbreviationWindow>();
            var page = App.Current.ServiceProvider.GetRequiredService<Pages.SimpleAbbreviationPage>();

            if (SelectedElement == BaseDictionaryElement.NoneElement)//(SelectedString == Properties.Resources.Other)
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
                var abbrev = savm.CurrentAbbrev ?? throw new NullReferenceException();
                App.Current.ServiceProvider.GetRequiredService<MainWindowViewModel>().TryWriteAbbreviation(abbrev, ShortForm);
            }

            w.Close();
        }
    }
}
