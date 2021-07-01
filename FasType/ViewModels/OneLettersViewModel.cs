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
        readonly static string _soloLetters;
        readonly IRepositoriesManager _repositories;
        ObservableCollection<BaseAbbreviation>? _oneLetters;

        public ObservableCollection<BaseAbbreviation>? OneLetters { get => _oneLetters; set => SetProperty(ref _oneLetters, value); }
        public Command<BaseAbbreviation> OpenAbbreviationPageCommand { get; }

        static OneLettersViewModel() => _soloLetters = @"befghikopqruvwxzéèçù";
        public OneLettersViewModel(IRepositoriesManager repositories)
        {
            _repositories = repositories;
            Init();

            OpenAbbreviationPageCommand = new(OpenAbbreviationPage, CanOpenAbbreviationPage);
        }

        //void Init(object sender, EventArgs e) => Init();
        void Init()
        {
            var ee = _soloLetters.Select(c => _repositories.Abbreviations[c.ToString()]).SelectMany((ba, i) => !ba.Any() ? Enumerable.Repeat(new SimpleAbbreviation($"{_soloLetters[i]}", "", 0, "", "", ""), 1) : ba);
            OneLetters = new(ee);
        }

        bool CanOpenAbbreviationPage() => !Windows.AbbreviationWindow.IsOpen;
        void OpenAbbreviationPage(BaseAbbreviation? abbrev)
        {
            _ = abbrev ?? throw new ArgumentNullException(nameof(abbrev));

            var aaw = App.Current.ServiceProvider.GetRequiredService<Windows.AbbreviationWindow>();
            var t = abbrev.GetModifyPageType();
            var p = (AbbreviationPage)App.Current.ServiceProvider.GetRequiredService(t);

            p.SetModifyAbbreviation(abbrev);

            aaw.Content = p;
            aaw.Closed += delegate { Init(); };
            aaw.Show();
        }
    }
}
