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
    public class SyllableAbbreviationViewModel : ObservableObject
    {
        readonly ILinguisticsStorage _storage;
        readonly SyllableAbbreviationRecord[] _arr;
        ObservableCollection<SyllableAbbreviation> _syllables;
        
        public Command<Window> SaveCommand { get; }
        public Command<SyllableAbbreviation> RemoveSyllableCommand { get; }
        public Command AddSyllableCommand { get; }
        public ObservableCollection<SyllableAbbreviation> Syllables { get => _syllables; set => SetProperty(ref _syllables, value); }

        public SyllableAbbreviationViewModel(ILinguisticsStorage storage)
        {
            _storage = storage;
            _arr = _storage.AbbreviationMethods.Cast<SyllableAbbreviationRecord>().ToArray();
            Syllables = new(_storage.AbbreviationMethods);//new(UserGrammar.SyllabesAbbreviations.Cast<SyllableAbbreviation>());

            AddSyllableCommand = new(AddSyllable);
            RemoveSyllableCommand = new(RemoveSyllable);
            SaveCommand = new(Save, CanSave);
        }

        void AddSyllable() => Syllables.Add(new(Guid.NewGuid(), string.Empty, string.Empty, SyllablePosition.None));//Syllables.Add(new(Guid.NewGuid(), "a", "a", SyllablePosition.In));
        void RemoveSyllable(SyllableAbbreviation sa)
        {
            var r = MessageBox.Show("Are you sure to delete this abbreviation method ?", Resources.Delete, MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes);
    
            if (r == MessageBoxResult.Yes)
                Syllables.Remove(sa);
        }

        bool CanSaveSyllable(SyllableAbbreviation sa) => !string.IsNullOrEmpty(sa.ShortForm) && !string.IsNullOrEmpty(sa.FullForm) && sa.Position != SyllablePosition.None;
        bool CanSave()
        {
            if (!Syllables.All(CanSaveSyllable))
                return false;

            if (Syllables.Count != _arr.Length)
                return true;

            if (Syllables.Any(sa => !_arr.Contains((SyllableAbbreviationRecord)sa)))
                return true;

            return false;
        }
        void Save(Window w)
        {
            _storage.AbbreviationMethods = Syllables;
            w.Close();
        }
    }
}
