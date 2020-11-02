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
        ObservableCollection<GrammarType> _plurals;

        public Command<Window> SaveCommand { get; }
        public Command<GrammarType> RemovePluralCommand { get; }
        public Command AddPluralCommand { get; }
        public ObservableCollection<GrammarType> Plurals { get => _plurals; set => SetProperty(ref _plurals, value); }

        public LinguisticsViewModel()
        {
            Plurals = new();

            SaveCommand = new(Save, CanSave);
            AddPluralCommand = new(AddPlural);
            RemovePluralCommand = new(RemovePlural);
            //SettingsToProperties();
        }

        void RemovePlural(GrammarType gt)
        {
            Plurals.Remove(gt);

            int i = 0;
            while (i < Plurals.Count)
                Plurals[i].Name = $"Plural {++i}";
        }

        void AddPlural() => Plurals.Add(new($"Plural {Plurals.Count + 1}", "k", GrammarPosition.Prefix));

        //void SettingsToProperties()
        //{
        //    UsesPT1 = _settings.PT1;
        //    IsPT1Postfix = _settings.PT1P;
        //    PT1Char = _settings.PT1C;
        //}

        //void PropertiesToSettings()
        //{
        //    _settings.PT1 = UsesPT1;
        //    _settings.PT1P = IsPT1Postfix;
        //    _settings.PT1C = PT1Char;
        //}

        bool CanSavePlural()
        {
            bool emptyReprs = Plurals.Select(gt => gt.Repr).Any(string.IsNullOrEmpty);
            if (emptyReprs)
                return false;

            bool emptyNames = Plurals.Select(gt => gt.Name).Any(string.IsNullOrEmpty);
            if (emptyNames)
                return false;

            bool duplicateNames = Plurals.Select(gt => gt.Name).Distinct().Count() < Plurals.Count;
            if (duplicateNames)
                return false;

            return true;
        }
        bool CanSave() => CanSavePlural();
        void Save(Window w)
        {
            //PropertiesToSettings();
            w.Close();
        }
    }
}
