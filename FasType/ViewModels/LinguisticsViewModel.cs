using FasType.Models;
using FasType.Models.Linguistics.Grammars;
using FasType.Properties;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Data;

namespace FasType.ViewModels
{
    public class LinguisticsViewModel: ObservableObject
    {
        private static GrammarTypeRecord genreRecord = new("Genre", "e", GrammarPosition.Prefix);
        private static GrammarTypeRecord pluralRecord = new("Plural", "k", GrammarPosition.Postfix);

        ObservableCollection<GrammarType> _plurals;

        public Command<Window> SaveCommand { get; }
        //public Command<GrammarType> RemovePluralCommand { get; }
        //public Command AddPluralCommand { get; }
        //public ObservableCollection<GrammarType> Plurals { get => _plurals; set => SetProperty(ref _plurals, value); }

        public GrammarType PluralContext { get; }
        public GrammarType GenreContext { get; }

        public LinguisticsViewModel()
        {
            Plurals = new();

            SaveCommand = new(Save, CanSave);
            //AddPluralCommand = new(AddPlural);
            //RemovePluralCommand = new(RemovePlural);
            //SettingsToProperties();

            GenreContext = (GrammarType)genreRecord;
            PluralContext = (GrammarType)pluralRecord;
        }

        //void RemovePlural(GrammarType gt)
        //{
        //    Plurals.Remove(gt);

        //    int i = 0;
        //    while (i < Plurals.Count)
        //        Plurals[i].Name = $"Plural {++i}";
        //}

        //void AddPlural() => Plurals.Add(new($"Plural {Plurals.Count + 1}", "k", GrammarPosition.Prefix));
        //bool CanSavePlural()
        //{
        //    bool emptyReprs = Plurals.Select(gt => gt.Repr).Any(string.IsNullOrEmpty);
        //    if (emptyReprs)
        //        return false;

        //    bool emptyNames = Plurals.Select(gt => gt.Name).Any(string.IsNullOrEmpty);
        //    if (emptyNames)
        //        return false;

        //    bool duplicateNames = Plurals.Select(gt => gt.Name).Distinct().Count() < Plurals.Count;
        //    if (duplicateNames)
        //        return false;

        //    return true;
        //}
        bool CanSavePlural()
        {
            bool emptyRepr = string.IsNullOrEmpty(PluralContext.Repr);
            if (emptyRepr)
                return false;

            bool emptyName = string.IsNullOrEmpty(PluralContext.Name);
            if (emptyName)
                return false;

            bool noChanges = pluralRecord == (GrammarTypeRecord)PluralContext;
            if (noChanges)
                return false;

            //bool duplicateNames = Plurals.Select(gt => gt.Name).Distinct().Count() < Plurals.Count;
            //if (duplicateNames)
            //    return false;

            return true;
        }
        bool CanSaveGenre()
        {
            bool emptyRepr = string.IsNullOrEmpty(GenreContext.Repr);
            if (emptyRepr)
                return false;

            bool emptyName = string.IsNullOrEmpty(GenreContext.Name);
            if (emptyName)
                return false;

            bool noChanges = genreRecord == (GrammarTypeRecord)GenreContext;
            if (noChanges)
                return false;

            return true;
        }
        bool CanSave() => CanSavePlural() || CanSaveGenre();

        void Save(Window w)
        {
            //PropertiesToSettings();
            pluralRecord = (GrammarTypeRecord)PluralContext;
            genreRecord = (GrammarTypeRecord)GenreContext;
            w.Close();
        }
    }
}
