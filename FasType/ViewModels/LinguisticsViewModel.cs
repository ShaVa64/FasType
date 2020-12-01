using FasType.Models;
using FasType.Models.Linguistics.Grammars;
using FasType.Properties;
using FasType.Storage;
using FasType.Windows;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Data;
using FasType.Services;
using Microsoft.Extensions.Configuration;
using FasType.Models.Linguistics;

namespace FasType.ViewModels
{
    public class LinguisticsViewModel: ObservableObject
    {
        //ObservableCollection<GrammarType> _plurals;
        //public Command<GrammarType> RemovePluralCommand { get; }
        //public Command AddPluralCommand { get; }
        //public ObservableCollection<GrammarType> Plurals { get => _plurals; set => SetProperty(ref _plurals, value); }
        readonly ILinguisticsStorage _storage;
        readonly IConfiguration _config;

        public Command<Window> SaveCommand { get; }
        public Command OpenSyllableCommand { get; }
        public Command ResetCommand { get; }
        public GrammarType PluralContext { get; private set; }
        public GrammarType GenderContext { get; private set; }
        public GrammarType GenderPluralContext { get; private set; }

        public LinguisticsViewModel(IConfiguration configuration, ILinguisticsStorage storage)
        {
            //Plurals = new();
            //AddPluralCommand = new(AddPlural);
            //RemovePluralCommand = new(RemovePlural);
            //SettingsToProperties();
            _storage = storage;
            _config = configuration;

            SaveCommand = new(Save, CanSave);
            OpenSyllableCommand = new(OpenSyllable);
            ResetCommand = new(Reset);

            GenderContext       = storage.GenderType;      //(GrammarType)UserGrammar.GenderRecord;
            PluralContext       = storage.PluralType;      //(GrammarType)UserGrammar.PluralRecord;
            GenderPluralContext = storage.GenderPluralType;//(GrammarType)UserGrammar.GenderPluralRecord;
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

        void Reset()
        {
            string path = _config.GetSection("Paths")["DefaultLinguistics"];

            var content = System.IO.File.ReadAllText(path);
            var dto = System.Text.Json.JsonSerializer.Deserialize<LinguisticsDTO>(content);

            GenderContext = dto.GenderType;
            PluralContext = dto.PluralType;
            GenderPluralContext = dto.GenderPluralType;

            _storage.AbbreviationMethods = dto.AbbreviationMethods;
            OnPropertyChanged(string.Empty);
        }

        void OpenSyllable()
        {
            var w = App.Current.ServiceProvider.GetRequiredService<SyllableAbbreviationWindow>();
            w.ShowDialog();
        }

        static bool CanSave(GrammarType context, GrammarTypeRecord record)
        {
            //if (string.IsNullOrEmpty(context.Repr)) //Empty Repr
            //    return false;

            //if (string.IsNullOrEmpty(context.Name)) //Empty Name
            //    return false;

            if ((GrammarTypeRecord)context == record) //No Changes
                return false;

            return true;
        }
        bool CanSavePlural() => CanSave(PluralContext, (GrammarTypeRecord)_storage.PluralType);
        bool CanSaveGenre() => CanSave(GenderContext, (GrammarTypeRecord)_storage.GenderType);
        bool CanSaveGenrePlural() => CanSave(GenderPluralContext, (GrammarTypeRecord)_storage.GenderPluralType);
        bool CanSave() => !string.IsNullOrEmpty(GenderContext.Repr)
                          && !string.IsNullOrEmpty(PluralContext.Repr)
                          && !string.IsNullOrEmpty(GenderPluralContext.Repr)
                          && (CanSavePlural() || CanSaveGenre() || CanSaveGenrePlural());
        void Save(Window w)
        {
            //PropertiesToSettings();
            _storage.PluralType = PluralContext;
            _storage.GenderType = GenderContext;
            _storage.GenderPluralType = GenderPluralContext;

            w.Close();
        }
    }
}
