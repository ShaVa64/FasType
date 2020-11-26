using FasType.Models;
using FasType.Models.Linguistics.Grammars;
using FasType.Properties;
using FasType.Storage;
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

        //ObservableCollection<GrammarType> _plurals;
        //public Command<GrammarType> RemovePluralCommand { get; }
        //public Command AddPluralCommand { get; }
        //public ObservableCollection<GrammarType> Plurals { get => _plurals; set => SetProperty(ref _plurals, value); }
        
        public Command<Window> SaveCommand { get; }
        public GrammarType PluralContext { get; }
        public GrammarType GenderContext { get; }
        public GrammarType PluralGenderContext { get; }

        public LinguisticsViewModel()
        {
            //Plurals = new();
            //AddPluralCommand = new(AddPlural);
            //RemovePluralCommand = new(RemovePlural);
            //SettingsToProperties();

            SaveCommand = new(Save, CanSave);

            GenderContext = (GrammarType)UserGrammar.GenderRecord;
            PluralContext = (GrammarType)UserGrammar.PluralRecord;
            PluralGenderContext = (GrammarType)UserGrammar.GenderPluralRecord;
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

        static bool CanSave(GrammarType context, GrammarTypeRecord record)
        {
            if (string.IsNullOrEmpty(context.Repr)) //Empty Repr
                return false;

            //if (string.IsNullOrEmpty(context.Name)) //Empty Name
            //    return false;

            if (record == (GrammarTypeRecord)context) //No Changes
                return false;

            return true;
        }
        bool CanSavePlural() => CanSave(PluralContext, UserGrammar.PluralRecord);
        bool CanSaveGenre() => CanSave(GenderContext, UserGrammar.GenderRecord);
        bool CanSaveGenrePlural() => CanSave(PluralGenderContext, UserGrammar.GenderPluralRecord);
        bool CanSave() => CanSavePlural() || CanSaveGenre() || CanSaveGenrePlural();
        void Save(Window w)
        {
            //PropertiesToSettings();
            UserGrammar.PluralRecord = (GrammarTypeRecord)PluralContext;
            UserGrammar.GenderRecord = (GrammarTypeRecord)GenderContext;
            UserGrammar.GenderPluralRecord = (GrammarTypeRecord)PluralGenderContext;
            w.Close();
        }
    }
}
