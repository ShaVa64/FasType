using FasType.Models;
using FasType.Core.Models.Linguistics;
using FasType.Properties;
using FasType.Windows;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Data;
using FasType.Core.Models;
using FasType.Core.Services;
using Microsoft.Extensions.Configuration;

namespace FasType.ViewModels
{
    public class LinguisticsViewModel : ObservableObject
    {
        //ObservableCollection<GrammarType> _plurals;
        //public Command<GrammarType> RemovePluralCommand { get; }
        //public Command AddPluralCommand { get; }
        //public ObservableCollection<GrammarType> Plurals { get => _plurals; set => SetProperty(ref _plurals, value); }
        readonly IRepositoriesManager _repositories;
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
                { nameof(GenderTypeContext), nameof(ILinguisticsRepository.GenderType) },
                { nameof(PluralTypeContext), nameof(ILinguisticsRepository.PluralType) },
                { nameof(GenderPluralTypeContext), nameof(ILinguisticsRepository.GenderPluralType) },
            };
            NoDupProperties = new string[] { nameof(GenderTypeContext), nameof(PluralTypeContext), nameof(GenderPluralTypeContext) };
        }

        public LinguisticsViewModel(IConfiguration configuration, IRepositoriesManager repositories)
        {
            //Plurals = new();
            //AddPluralCommand = new(AddPlural);
            //RemovePluralCommand = new(RemovePlural);
            //SettingsToProperties();
            _repositories = repositories;
            _config = configuration;

            SaveCommand = new(Save, CanSave);
            OpenSyllableCommand = new(OpenSyllable, CanOpenSyllable);
            ResetCommand = new(Reset, CanReset);
            OpenOneLettersCommand = new(OpenOneLetters, CanOpenOneLetters);

            GenderTypeContext = _repositories.Linguistics.GenderType;
            PluralTypeContext = _repositories.Linguistics.PluralType;
            GenderPluralTypeContext = _repositories.Linguistics.GenderPluralType;
        }

        bool CanOpenOneLetters() => !OneLettersWindow.IsOpen;
        void OpenOneLetters()
        {
            var w = App.Current.ServiceProvider.GetRequiredService<OneLettersWindow>();

            w.Show();
        }

        bool CanReset() => !AbbreviationMethodsWindow.IsOpen;
        void Reset()
        {
            string path = _config.GetSection("Paths")["DefaultLinguistics"];

            var content = System.IO.File.ReadAllText(path);
            var dto = System.Text.Json.JsonSerializer.Deserialize<LinguisticsDTO>(content) ?? throw new NullReferenceException();

            GenderTypeContext = dto.GenderType;
            PluralTypeContext = dto.PluralType;
            GenderPluralTypeContext = dto.GenderPluralType;

            _repositories.Linguistics.AbbreviationMethods = dto.AbbreviationMethods;
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

            if (context == record)
                return false;

            return true;
        }
        bool EmptyRepr() => GetType().GetProperties().Where(pi => pi.PropertyType == typeof(GrammarType)).Select(pi => (pi.GetValue(this) as GrammarType ?? throw new NullReferenceException()).Repr).ToList().Any(string.IsNullOrEmpty);
        bool CanSaveGrammarType(string propName) => CanSave(typeof(LinguisticsViewModel).GetProperty(propName)?.GetValue(this) as GrammarType,
            typeof(ILinguisticsRepository).GetProperty(PropertiesContextPair[propName])?.GetValue(_repositories.Linguistics) as GrammarType);

        bool NoDup() => NoDupProperties.Select(s => (typeof(LinguisticsViewModel).GetProperty(s)?.GetValue(this) as GrammarType)?.Repr).Distinct().Count() == NoDupProperties.Length;

        bool CanSave() => !EmptyRepr() && NoDup() && PropertiesContextPair.Keys.ToList().Any(CanSaveGrammarType);
        void Save(Window? w)
        {
            _ = w ?? throw new NullReferenceException();
            _repositories.Linguistics.PluralType = PluralTypeContext;
            _repositories.Linguistics.GenderType = GenderTypeContext;
            _repositories.Linguistics.GenderPluralType = GenderPluralTypeContext;

            w.Close();
        }
    }
}
