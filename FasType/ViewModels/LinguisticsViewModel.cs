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
        static readonly Dictionary<string, string> PropertiesContextPair;
        static readonly string[] NoDupProperties;

        public Command<Window> SaveCommand { get; }
        public Command OpenSyllableCommand { get; }
        public Command ResetCommand { get; }
        public Command OpenOneLettersCommand { get; }

        public GrammarType GenderTypeContext { get; private set; }
        public GrammarType PluralTypeContext { get; private set; }
        public GrammarType GenderPluralTypeContext { get; private set; }

        static LinguisticsViewModel()
        {
            PropertiesContextPair = new()
            {
                { nameof(GenderTypeContext), nameof(ILinguisticsStorage.GenderType) },
                { nameof(PluralTypeContext), nameof(ILinguisticsStorage.PluralType) },
                { nameof(GenderPluralTypeContext), nameof(ILinguisticsStorage.GenderPluralType) },
            };
            NoDupProperties = new string[] { nameof(GenderTypeContext), nameof(PluralTypeContext), nameof(GenderPluralTypeContext) };
        }

        public LinguisticsViewModel(IConfiguration configuration, ILinguisticsStorage storage)
        {
            //Plurals = new();
            //AddPluralCommand = new(AddPlural);
            //RemovePluralCommand = new(RemovePlural);
            //SettingsToProperties();
            _storage = storage;
            _config = configuration;

            SaveCommand = new(Save, CanSave);
            OpenSyllableCommand = new(OpenSyllable, CanOpenSyllable);
            ResetCommand = new(Reset, CanReset);
            OpenOneLettersCommand = new(OpenOneLetters, CanOpenOneLetters);

            GenderTypeContext       = storage.GenderType;      //(GrammarType)UserGrammar.GenderRecord;
            PluralTypeContext       = storage.PluralType;      //(GrammarType)UserGrammar.PluralRecord;
            GenderPluralTypeContext = storage.GenderPluralType;//(GrammarType)UserGrammar.GenderPluralRecord;
        }

        bool CanOpenOneLetters() => !OneLettersWindow.IsOpen;
        void OpenOneLetters()
        {
            var w = App.Current.ServiceProvider.GetRequiredService<OneLettersWindow>();

            w.Show();
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

        bool CanReset() => !AbbreviationMethodsWindow.IsOpen;
        void Reset()
        {
            string path = _config.GetSection("Paths")["DefaultLinguistics"];

            var content = System.IO.File.ReadAllText(path);
            var dto = System.Text.Json.JsonSerializer.Deserialize<LinguisticsDTO>(content) ?? throw new NullReferenceException();

            GenderTypeContext = dto.GenderType;
            PluralTypeContext = dto.PluralType;
            GenderPluralTypeContext = dto.GenderPluralType;

            _storage.AbbreviationMethods = dto.AbbreviationMethods;
            OnPropertyChanged(string.Empty);
        }

        bool CanOpenSyllable() => !AbbreviationMethodsWindow.IsOpen;
        void OpenSyllable()
        {
            var w = App.Current.ServiceProvider.GetRequiredService<AbbreviationMethodsWindow>();

            w.Show();
        }

        static bool CanSave(GrammarType? context, GrammarType? record)
        {
            _ = context ?? throw new NullReferenceException();
            _ = record ?? throw new NullReferenceException();
            //if (string.IsNullOrEmpty(context.Repr)) //Empty Repr
            //    return false;

            //if (string.IsNullOrEmpty(context.Name)) //Empty Name
            //    return false;

            if (((GrammarTypeRecord)context) == ((GrammarTypeRecord)record)) //No Changes
                return false;

            return true;
        }
        bool EmptyRepr() => GetType().GetProperties().Where(pi => pi.PropertyType == typeof(GrammarType)).Select(pi => (pi.GetValue(this) as GrammarType ?? throw new NullReferenceException()).Repr).ToList().Any(string.IsNullOrEmpty);
        //string.IsNullOrEmpty(GenderTypeContext.Repr)
        //                    || string.IsNullOrEmpty(PluralTypeContext.Repr)
        //                    || string.IsNullOrEmpty(GenderPluralTypeContext.Repr)
        //                    || string.IsNullOrEmpty(GenderCompletionContext.Repr)
        //                    || string.IsNullOrEmpty(PluralCompletionContext.Repr)
        //                    || string.IsNullOrEmpty(GenderPluralCompletionContext.Repr);
        bool CanSaveGrammarType(string propName) => CanSave(typeof(LinguisticsViewModel).GetProperty(propName)?.GetValue(this) as GrammarType, 
            typeof(ILinguisticsStorage).GetProperty(PropertiesContextPair[propName])?.GetValue(_storage) as GrammarType);

        bool NoDup() => NoDupProperties.Select(s => (typeof(LinguisticsViewModel).GetProperty(s)?.GetValue(this) as GrammarType)?.Repr).Distinct().Count() == NoDupProperties.Length;//.ToList()
        //bool CanSavePluralType() => CanSave(PluralTypeContext, (GrammarTypeRecord)_storage.PluralType);
        //bool CanSaveGenderType() => CanSave(GenderTypeContext, (GrammarTypeRecord)_storage.GenderType);
        //bool CanSaveGenrePluralType() => CanSave(GenderPluralTypeContext, (GrammarTypeRecord)_storage.GenderPluralType);
        //bool CanSavePluralCompletion() => CanSave(PluralCompletionContext, (GrammarTypeRecord)_storage.PluralCompletion);
        //bool CanSaveGenderCompletion() => CanSave(GenderCompletionContext, (GrammarTypeRecord)_storage.GenderCompletion);
        //bool CanSaveGenrePluralCompletion() => CanSave(GenderPluralCompletionContext, (GrammarTypeRecord)_storage.GenderPluralCompletion);
        bool CanSave() => !EmptyRepr() && NoDup() && PropertiesContextPair.Keys.ToList().Any(CanSaveGrammarType);//(CanSavePluralType() || CanSaveGenderType() || CanSaveGenrePluralType() || CanSavePluralCompletion() || CanSaveGenderCompletion() || CanSaveGenrePluralCompletion());
        void Save(Window? w)
        {
            _ = w ?? throw new NullReferenceException();
            //PropertiesToSettings();
            _storage.PluralType = PluralTypeContext;
            _storage.GenderType = GenderTypeContext;
            _storage.GenderPluralType = GenderPluralTypeContext;

            w.Close();
        }
    }
}
