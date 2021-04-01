using FasType.Models;
using FasType.Utils;
using FasType.Models.Linguistics;
using FasType.Properties;
using FasType.Storage;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using FasType.Services;

namespace FasType.ViewModels
{
    public class AbbreviationMethodsViewModel : ObservableObject
    {
        readonly ILinguisticsStorage _storage;
        readonly AbbreviationMethodRecord[] _arr;
        ObservableCollection<AbbreviationMethod> _syllables;

        public string Title => Resources.AbbreviationMethod + $"  ({Syllables.Count})";
        public Command<Window> SaveCommand { get; }
        public Command<AbbreviationMethod> RemoveSyllableCommand { get; }
        public Command AddSyllableCommand { get; }
        public ObservableCollection<AbbreviationMethod> Syllables { get => _syllables; set => SetProperty(ref _syllables, value); }

        public AbbreviationMethodsViewModel(ILinguisticsStorage storage)
        {
            _storage = storage;
            _arr = _storage.AbbreviationMethods.Cast<AbbreviationMethodRecord>()/*.OrderBy(amr => amr.)*/.ToArray();
            Syllables = new(_storage.AbbreviationMethods);//new(UserGrammar.SyllabesAbbreviations.Cast<SyllableAbbreviation>());
            _ = _syllables ?? throw new NullReferenceException();

            AddSyllableCommand = new(AddSyllable);
            RemoveSyllableCommand = new(RemoveSyllable);
            SaveCommand = new(Save, CanSave);
        }

        void AddSyllable()
        {
            Syllables.Add(new(Guid.NewGuid(), string.Empty, string.Empty, SyllablePosition.None));//Syllables.Add(new(Guid.NewGuid(), "a", "a", SyllablePosition.In));
            OnPropertyChanged(nameof(Title));
        }

        void RemoveSyllable(AbbreviationMethod? sa)
        {
            _ = sa ?? throw new NullReferenceException();
            var r = MessageBox.Show(DialogResources.DeleteMethodDialog, Resources.Delete, MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes);
    
            if (r == MessageBoxResult.Yes)
            {
                Syllables.Remove(sa);
                OnPropertyChanged(nameof(Title));
            }
        }

        bool CanSaveSyllable(AbbreviationMethod sa) => !string.IsNullOrEmpty(sa.ShortForm) && !string.IsNullOrEmpty(sa.FullForm) && sa.Position != SyllablePosition.None;
        bool CanSave()
        {
            if (!Syllables.All(CanSaveSyllable))
                return false;

            if (Syllables.Count != _arr.Length)
                return true;

            if (Syllables.Any(sa => !_arr.Contains((AbbreviationMethodRecord)sa)))
                return true;

            return false;
        }
        void Save(Window? w)
        {
            _ = w ?? throw new NullReferenceException();
            _storage.AbbreviationMethods = Syllables;
            w.Close();
        }
    }
}
