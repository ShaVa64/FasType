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
using System.Windows.Data;
using System.Windows.Input;

namespace FasType.ViewModels
{
    public class SeeAllViewModel : ObservableObject
    {
        string _queryString;
        FormOrderBy _sortBy;
        readonly IAbbreviationStorage _storage;
        List<BaseAbbreviation> _allAbbreviations;

        public string Title => Resources.AllAbbrevs + $"  ({Count})";

        public FormOrderBy OrderBy
        {
            get => _sortBy;
            set
            {
                if (SetProperty(ref _sortBy, value))
                    OrderAndFilterAbbreviations();
            }
        }

        public string QueryString
        {
            get => _queryString;
            set
            {
                if (SetProperty(ref _queryString, value))
                    OrderAndFilterAbbreviations();
            }
        }

        public int Count => AllAbbreviations.Count;
        public List<BaseAbbreviation> AllAbbreviations
        {
            get => _allAbbreviations;
            private set
            {
                SetProperty(ref _allAbbreviations, value);
                OnPropertyChanged(nameof(Count));
                OnPropertyChanged(nameof(Title));
            }
        }

        public Command<BaseAbbreviation> RemoveCommand { get; }

        public SeeAllViewModel(IAbbreviationStorage storage)
        {
            _storage = storage;

            RemoveCommand = new(Remove, CanRemove);
            _queryString = "";
            OrderAbbreviations();
            //AllAbbreviations = _storage.Take(2).ToList();
        }

        void OrderAndFilterAbbreviations() => AllAbbreviations = (OrderBy switch
        {
            FormOrderBy.FullForm => _storage.Where(a => a.FullForm.Contains(QueryString)).OrderBy(a => a.FullForm),
            FormOrderBy.ShortForm => _storage.Where(a => a.ShortForm.Contains(QueryString)).OrderBy(a => a.ShortForm),
            _ => throw new NotImplementedException()
        }).ToList();
        void OrderAbbreviations() => AllAbbreviations = (OrderBy switch
        {
            FormOrderBy.FullForm => _storage.OrderBy(a => a.FullForm),
            FormOrderBy.ShortForm => _storage.OrderBy(a => a.ShortForm),
            _ => throw new NotImplementedException()
        }).ToList();
        void FilterAbbreviations() =>  AllAbbreviations = (OrderBy switch
        {
            FormOrderBy.FullForm => _storage.Where(a => a.FullForm.Contains(QueryString)),
            FormOrderBy.ShortForm => _storage.Where(a => a.ShortForm.Contains(QueryString)),
            _ => throw new NotImplementedException()
        }).ToList();

        bool CanRemove() => true;
        void Remove(BaseAbbreviation abbrev)
        {
            var message = string.Format(DialogResources.DeleteDialogFormat, Environment.NewLine, abbrev.ElementaryRepresentation);
            var res = MessageBox.Show(message, Resources.Delete, MessageBoxButton.OKCancel, MessageBoxImage.Question);
            if (res == MessageBoxResult.Cancel)
                return;

            _storage.Remove(abbrev);
            OrderAndFilterAbbreviations();
        }
        public enum FormOrderBy
        {
            ShortForm,
            FullForm
        }
    }
}
