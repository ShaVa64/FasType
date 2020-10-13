using FasType.Models;
using FasType.Properties;
using FasType.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace FasType.ViewModels
{
    public class SeeAllViewModel : BaseViewModel
    {
        readonly IDataStorage _storage;
        IList<IAbbreviation> _allAbbreviations;

        public int Count => AllAbbreviations.Count;
        public IList<IAbbreviation> AllAbbreviations
        {
            get => _allAbbreviations;
            private set
            {
                SetProperty(ref _allAbbreviations, value);
                OnPropertyChanged(nameof(Count));
            }
        }

        public RoutedCommand RemoveCommand { get; }

        public SeeAllViewModel(IDataStorage storage)
        {
            _storage = storage;

            RemoveCommand = new RoutedCommand("Remove", typeof(SeeAllViewModel));
            AllAbbreviations = _storage.ToList();
            //AllAbbreviations = _storage.Take(2).ToList();
        }

        public void CanRemove(object sender, CanExecuteRoutedEventArgs e) => e.CanExecute = true;
        public void Remove(object sender, ExecutedRoutedEventArgs e)
        {
            var abbrev = e.Parameter as IAbbreviation;

            var message = string.Format(Resources.DeleteDIalogFormat, Environment.NewLine, abbrev.ElementaryRepresentation);
            var res = MessageBox.Show(message, Resources.Delete, MessageBoxButton.OKCancel, MessageBoxImage.Question);
            if (res == MessageBoxResult.Cancel)
                return;

            _storage.Remove(abbrev);
            AllAbbreviations = _storage.ToList();
        }
    }

    public class SeeAllSelector : DataTemplateSelector
    {
        public DataTemplate First { get; set; }
        public DataTemplate Default { get; set; }
        public DataTemplate Last { get; set; }
        public DataTemplate Only { get; set; }

        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            int altIndex = ItemsControl.GetAlternationIndex(container);

            var ic = ItemsControl.ItemsControlFromItemContainer(container);
            int altCount = ic.AlternationCount;

            if (altCount == 1)
                return Only;
            if (altIndex == 0)
                return First;
            if (altIndex == altCount - 1)
                return Last;
            
            return Default;
            //return base.SelectTemplate(item, container);
        }
    }
}
