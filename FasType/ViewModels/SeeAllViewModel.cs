using FasType.Models;
using FasType.Core.Models;
using FasType.Core.Models.Abbreviations;
using FasType.Properties;
using FasType.Core.Services;
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
        readonly IRepositoriesManager _repositories;
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

        public ILinguisticsRepository Linguistics => _repositories.Linguistics;

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

        public SeeAllViewModel(IRepositoriesManager repositories)
        {
            _repositories = repositories;

            RemoveCommand = new(Remove, CanRemove);
            _queryString = "";
            OrderAbbreviations();
            _ = _allAbbreviations ?? throw new NullReferenceException();
        }

        void OrderAndFilterAbbreviations() => AllAbbreviations = (OrderBy switch
        {
            FormOrderBy.FullForm => _repositories.Abbreviations.Where(a => a.FullForm.Contains(QueryString)).OrderBy(a => a.FullForm).ThenBy(a => a.ShortForm),
            FormOrderBy.ShortForm => _repositories.Abbreviations.Where(a => a.ShortForm.Contains(QueryString)).OrderBy(a => a.ShortForm).ThenBy(a => a.FullForm),
            _ => throw new NotImplementedException()
        }).ToList();
        void OrderAbbreviations() => AllAbbreviations = (OrderBy switch
        {
            FormOrderBy.FullForm => _repositories.Abbreviations.GetAll().OrderBy(a => a.FullForm).ThenBy(a => a.ShortForm),
            FormOrderBy.ShortForm => _repositories.Abbreviations.GetAll().OrderBy(a => a.ShortForm).ThenBy(a => a.FullForm),
            _ => throw new NotImplementedException()
        }).ToList();
        void FilterAbbreviations() =>  AllAbbreviations = (OrderBy switch
        {
            FormOrderBy.FullForm => _repositories.Abbreviations.Where(a => a.FullForm.Contains(QueryString)),
            FormOrderBy.ShortForm => _repositories.Abbreviations.Where(a => a.ShortForm.Contains(QueryString)),
            _ => throw new NotImplementedException()
        }).ToList();

        bool CanRemove() => true;
        void Remove(BaseAbbreviation? abbrev)
        {
            _ = abbrev ?? throw new NullReferenceException();
            var message = string.Format(DialogResources.DeleteDialogFormat, Environment.NewLine, abbrev.ElementaryRepresentation);
            var res = MessageBox.Show(message, Resources.Delete, MessageBoxButton.OKCancel, MessageBoxImage.Question);
            if (res == MessageBoxResult.Cancel)
                return;

            _repositories.Abbreviations.Remove(abbrev);
            _repositories.Abbreviations.SaveChanges();
            OrderAndFilterAbbreviations();
        }
        public enum FormOrderBy
        {
            ShortForm,
            FullForm
        }
    }
}
