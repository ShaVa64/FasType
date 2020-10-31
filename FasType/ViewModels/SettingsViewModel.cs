using FasType.Models;
using FasType.Properties;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Data;

namespace FasType.ViewModels
{
    public class SettingsViewModel: BaseViewModel
    {
        readonly Settings _settings;

        bool _usesPT1, _isPT1Postfix;
        string _pT1Char;

        public bool UsesPT1 { get => _usesPT1; set => SetProperty(ref _usesPT1, value); }
        public bool IsPT1Postfix { get => _isPT1Postfix; set => SetProperty(ref _isPT1Postfix, value); }
        public string PT1Char { get => _pT1Char; set => SetProperty(ref _pT1Char, value); }

        public Command<Window> SaveCommand { get; }

        public SettingsViewModel()
        {
            _settings = Settings.Default;
            SaveCommand = new(Save, CanSave);

            SettingsToProperties();
        }

        void SettingsToProperties()
        {
            UsesPT1 = _settings.PT1;
            IsPT1Postfix = _settings.PT1P;
            PT1Char = _settings.PT1C;
        }

        void PropertiesToSettings()
        {
            _settings.PT1 = UsesPT1;
            _settings.PT1P = IsPT1Postfix;
            _settings.PT1C = PT1Char;
        }

        bool CanSavePlural()
        {
            bool pluralTypeChange = UsesPT1 != _settings.PT1;
            bool pluralSettingsChange = UsesPT1 == true && _settings.PT1 == true && (IsPT1Postfix != _settings.PT1P || PT1Char != _settings.PT1C);

            return pluralTypeChange || pluralSettingsChange;
        }
        bool CanSave() => CanSavePlural();
        void Save(Window w)
        {
            PropertiesToSettings();
            _settings.Save();
            w.Close();
        }
    }
}
