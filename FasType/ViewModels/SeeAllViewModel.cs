using FasType.Models;
using FasType.Models.Abbreviations;
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
        string _queryString;
        FormOrderBy _sortBy;
        readonly IDataStorage _storage;
        IList<IAbbreviation> _allAbbreviations;

        public FormOrderBy OrderBy
        {
            get => _sortBy;
            set
            {
                SetProperty(ref _sortBy, value);
                AllAbbreviations = (OrderBy switch
                {
                    FormOrderBy.FullForm => AllAbbreviations.OrderBy(a => a.FullForm),
                    FormOrderBy.ShortForm => AllAbbreviations.OrderBy(a => a.ShortForm),
                    _ => throw new NotImplementedException()
                }).ToList();
            }
        }

        public string QueryString
        {
            get => _queryString;
            set
            {
                if (SetProperty(ref _queryString, value.TrimStart()))
                    AllAbbreviations = (OrderBy switch
                    {
                        FormOrderBy.FullForm => _storage.Where(a => a.FullForm.StartsWith(QueryString)),
                        FormOrderBy.ShortForm => _storage.Where(a => a.ShortForm.StartsWith(QueryString)),
                        _ => throw new NotImplementedException()
                    }).ToList();
            }
        }

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

        public Command<IAbbreviation> RemoveCommand { get; }

        public SeeAllViewModel(IDataStorage storage)
        {
            _storage = storage;

            RemoveCommand = new(Remove, CanRemove);
            AllAbbreviations = _storage.ToList();
            //AllAbbreviations = _storage.Take(2).ToList();
        }

        bool CanRemove() => true;
        void Remove(IAbbreviation abbrev)
        {
            var message = string.Format(Resources.DeleteDialogFormat, Environment.NewLine, abbrev.ElementaryRepresentation);
            var res = MessageBox.Show(message, Resources.Delete, MessageBoxButton.OKCancel, MessageBoxImage.Question);
            if (res == MessageBoxResult.Cancel)
                return;

            _storage.Remove(abbrev);
            AllAbbreviations = _storage.ToList();
        }

        public enum FormOrderBy
        {
            ShortForm,
            FullForm
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
        }
    }
}
