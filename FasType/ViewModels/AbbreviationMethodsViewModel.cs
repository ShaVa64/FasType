using FasType.Core.Models;
using FasType.Models;
using FasType.Utils;
using FasType.Core.Models.Linguistics;
using FasType.Properties;
using FasType.Core.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace FasType.ViewModels
{
    public class AbbreviationMethodsViewModel : ObservableObject
    {
        readonly IRepositoriesManager _repositories;
        readonly AbbreviationMethod[] _arr;
        ObservableCollection<AbbreviationMethod> _syllables;

        public string Title => Resources.AbbreviationMethod + $"  ({AbbreviationMethods.Count})";
        public Command<Window> SaveCommand { get; }
        public Command<AbbreviationMethod> RemoveSyllableCommand { get; }
        public Command AddSyllableCommand { get; }
        public ObservableCollection<AbbreviationMethod> AbbreviationMethods { get => _syllables; set => SetProperty(ref _syllables, value); }

        public AbbreviationMethodsViewModel(IRepositoriesManager repositories)
        {
            _repositories = repositories;
            _arr = _repositories.Linguistics.AbbreviationMethods.OrderBy(amr => amr.ShortForm).ThenBy(amr => amr.FullForm).ToArray();
            AbbreviationMethods = new(_repositories.Linguistics.AbbreviationMethods);
            _ = _syllables ?? throw new NullReferenceException();

            AddSyllableCommand = new(AddSyllable);
            RemoveSyllableCommand = new(RemoveSyllable);
            SaveCommand = new(Save, CanSave);
        }

        void AddSyllable()
        {
            AbbreviationMethods.Add(new(Guid.NewGuid(), string.Empty, string.Empty, SyllablePosition.None));//Syllables.Add(new(Guid.NewGuid(), "a", "a", SyllablePosition.In));
            OnPropertyChanged(nameof(Title));
        }

        void RemoveSyllable(AbbreviationMethod? sa)
        {
            _ = sa ?? throw new NullReferenceException();
            var r = MessageBox.Show(DialogResources.DeleteMethodDialog, Resources.Delete, MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes);
    
            if (r == MessageBoxResult.Yes)
            {
                AbbreviationMethods.Remove(sa);
                OnPropertyChanged(nameof(Title));
            }
        }

        bool CanSaveSyllable(AbbreviationMethod sa) => !string.IsNullOrEmpty(sa.ShortForm) && !string.IsNullOrEmpty(sa.FullForm) && sa.Position != SyllablePosition.None;
        bool CanSave()
        {
            if (!AbbreviationMethods.All(CanSaveSyllable))
                return false;

            if (AbbreviationMethods.Count != _arr.Length)
                return true;

            if (AbbreviationMethods.Any(sa => !_arr.Contains(sa)))
                return true;

            return false;
        }
        void Save(Window? w)
        {
            _ = w ?? throw new NullReferenceException();
            _repositories.Linguistics.AbbreviationMethods = AbbreviationMethods;
            _repositories.Linguistics.SaveChanges();
            w.Close();
        }
    }
}
