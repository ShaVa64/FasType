using FasType.Models;
using FasType.Models.Dictionary;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace FasType.ViewModels
{
    public class PopupViewModel : ObservableObject
    {
        string? _shortForm/*, _selectedString*/;
        ObservableCollection<BaseDictionaryElement>? _collection;
        BaseDictionaryElement[]? _elements;
        BaseDictionaryElement? _selectedElement;

        static Services.ILinguisticsStorage Linguistics => App.Current.ServiceProvider.GetRequiredService<Services.ILinguisticsStorage>();
        static Services.IDictionaryStorage Dictionary => App.Current.ServiceProvider.GetRequiredService<Services.IDictionaryStorage>();

        public ObservableCollection<BaseDictionaryElement>? Collection { get => _collection; set => SetProperty(ref _collection, value); }
        public string? ShortForm { get => _shortForm; set => SetProperty(ref _shortForm, value); }
        public Command<Window> CreateCommand { get; }
        public BaseDictionaryElement? SelectedElement { get => _selectedElement; set => SetProperty(ref _selectedElement, value); }
        //public string SelectedString { get => _selectedString; set => SetProperty(ref _selectedString, value); }


        public PopupViewModel()
        {
            CreateCommand = new(Create, CanCreate);
        }

        public void SearchForWord(string currentWord)
        {
            ShortForm = currentWord;
            var l = Linguistics.Words(currentWord).ToList();

            _elements = l.SelectMany(s => Dictionary.GetElements(s, 1) ?? Array.Empty<BaseDictionaryElement>()).Distinct().ToArray();

            Collection = new ObservableCollection<BaseDictionaryElement>(_elements/*.Select(e => e.FullForm)*/)
            {
                BaseDictionaryElement.OtherElement,
                BaseDictionaryElement.NoneElement
                //Properties.Resources.Other,
                //Properties.Resources.None
            };
            SelectedElement = Collection[0];
            //SelectedString = Collection[0];
        }

        bool CanCreate() => SelectedElement != null;//!string.IsNullOrEmpty(SelectedString);
        void Create(Window? w)
        {
            _ = w ?? throw new NullReferenceException();
            if (SelectedElement == BaseDictionaryElement.NoneElement)//(SelectedString == Properties.Resources.None)
            {
                var r = MessageBox.Show(Properties.DialogResources.AddDictionary, Properties.Resources.Dictionary, MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (r == MessageBoxResult.Yes)
                    Dictionary.Add(new SimpleDictionaryElement(ShortForm ?? throw new NullReferenceException(), string.Empty, string.Empty, string.Empty));
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

            window.Content = page;
            window.Show();

            w.Close();
        }
    }
}
