using FasType.Models;
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
        string _shortForm, _selectedString;
        ObservableCollection<string> _collection;
        Models.Dictionary.BaseDictionaryElement[] _elements;

        static Services.ILinguisticsStorage Linguistics => App.Current.ServiceProvider.GetRequiredService<Services.ILinguisticsStorage>();
        static Services.IDictionaryStorage Dictionary => App.Current.ServiceProvider.GetRequiredService<Services.IDictionaryStorage>();

        public ObservableCollection<string> Collection { get => _collection; set => SetProperty(ref _collection, value); }
        public string ShortForm { get => _shortForm; set => SetProperty(ref _shortForm, value); }
        public Command<Window> CreateCommand { get; }
        public string SelectedString { get => _selectedString; set => SetProperty(ref _selectedString, value); }


        public PopupViewModel()
        {
            CreateCommand = new(Create, CanCreate);
        }

        public void SearchForWord(string currentWord)
        {
            ShortForm = currentWord;
            var l = Linguistics.Words(currentWord).ToList();

            _elements = l.Select(s => Dictionary.GetElement(s)).Where(e => e != null).ToArray();

            Collection = new ObservableCollection<string>(_elements.Select(e => e.FullForm))
            {
                Properties.Resources.Other,
                Properties.Resources.None
            };
            SelectedString = Collection[0];
        }

        bool CanCreate() => !string.IsNullOrEmpty(SelectedString);
        void Create(Window w)
        {
            if (SelectedString == Properties.Resources.None)
            {
                var r = MessageBox.Show(Properties.DialogResources.AddDictionary, Properties.Resources.Dictionary, MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (r == MessageBoxResult.Yes)
                    Dictionary.Add(new Models.Dictionary.SimpleDictionaryElement(ShortForm, string.Empty, string.Empty, string.Empty));
                w.Close();
                return;
            }

            var window = App.Current.ServiceProvider.GetRequiredService<Windows.AbbreviationWindow>();
            var page = App.Current.ServiceProvider.GetRequiredService<Pages.SimpleAbbreviationPage>();

            if (SelectedString == Properties.Resources.Other)
            {
                page.SetNewAbbreviation(ShortForm, "", null);
            }
            else
            {
                page.SetNewAbbreviation(ShortForm, SelectedString, Dictionary.GetElement(SelectedString)?.Others);
            }

            window.Content = page;
            window.Show();

            w.Close();
        }
    }
}
