using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using FasType.Models;
using FasType.Pages;
using FasType.Core.Services;
using Microsoft.Extensions.DependencyInjection;
using FasType.Utils;
using FasType.Core.Models.Abbreviations;
using FasType.Core.Models;

namespace FasType.ViewModels
{
    public class OneLettersViewModel : ObservableObject
    {
        private readonly static string _soloLetters;
        private readonly IRepositoriesManager _repositories;
        private ObservableCollection<OneLettersAbbreviationViewModel> _oneLetters;

        public ObservableCollection<OneLettersAbbreviationViewModel> OneLetters { get => _oneLetters; set => SetProperty(ref _oneLetters, value); }
        public Command<BaseAbbreviation> OpenAbbreviationPageCommand { get; }

        static OneLettersViewModel() => _soloLetters = @"befghikopqruvwxzéèçù";
        public OneLettersViewModel(IRepositoriesManager repositories)
        {
            _repositories = repositories;
            _oneLetters = new();
            OpenAbbreviationPageCommand = new(OpenAbbreviationPage, CanOpenAbbreviationPage);

            Init();
        }

        private void Init()
        {
            var ee = _soloLetters.Select(c => _repositories.Abbreviations[c.ToString()]).SelectMany((ba, i) => !ba.Any() ? Enumerable.Repeat(new SimpleAbbreviation($"{_soloLetters[i]}", "", 0, "", "", ""), 1) : ba);

            var vms = ee.Select(ba => new OneLettersAbbreviationViewModel(_repositories, ba, OpenAbbreviationPageCommand));
            OneLetters = new(vms);
        }

        private bool CanOpenAbbreviationPage() => !Windows.AbbreviationWindow.IsOpen;
        private void OpenAbbreviationPage(BaseAbbreviation? abbrev)
        {
            _ = abbrev ?? throw new ArgumentNullException(nameof(abbrev));

            var aaw = App.Current.ServiceProvider.GetRequiredService<Windows.AbbreviationWindow>();
            var t = abbrev.GetModifyPageType();
            var p = (AbbreviationPage)App.Current.ServiceProvider.GetRequiredService(t);

            if (abbrev.FullForm != string.Empty)
            {
                p.SetModifyAbbreviation(abbrev);
            }
            else
            {
                p.SetNewAbbreviation(abbrev.ShortForm, string.Empty);
            }

            aaw.Content = p;
            aaw.Closed += delegate
            {
                _repositories.Reload();
                Init();
            };
            aaw.Show();
        }
    }
}
