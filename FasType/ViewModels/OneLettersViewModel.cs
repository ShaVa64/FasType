using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using FasType.Models;
using FasType.Models.Abbreviations;
using FasType.Pages;
using FasType.Services;
using Microsoft.Extensions.DependencyInjection;
using FasType.Utils;

namespace FasType.ViewModels
{
    public class OneLettersViewModel : ObservableObject
    {
        readonly static string _soloLetters;
        ObservableCollection<BaseAbbreviation> _oneLetters;

        static IAbbreviationStorage Storage => App.Current.ServiceProvider.GetRequiredService<IAbbreviationStorage>();
        public ObservableCollection<BaseAbbreviation> OneLetters { get => _oneLetters; set => SetProperty(ref _oneLetters, value); }
        public Command<BaseAbbreviation> OpenAbbreviationPageCommand { get; }

        static OneLettersViewModel() => _soloLetters = @"befghikopqruvwxzéèçù";
        public OneLettersViewModel()
        {
            Init();

            OpenAbbreviationPageCommand = new(OpenAbbreviationPage, CanOpenAbbreviationPage);
        }

        void Init(object sender, EventArgs e) => Init();
        void Init()
        {
            var ee = _soloLetters.Select(c => Storage[c.ToString()]).SelectMany(ba => ba);
            OneLetters = new(ee);
        }

        bool CanOpenAbbreviationPage() => !Windows.AbbreviationWindow.IsOpen;
        void OpenAbbreviationPage(BaseAbbreviation abbrev)
        {
            var aaw = App.Current.ServiceProvider.GetRequiredService<Windows.AbbreviationWindow>();
            var t = abbrev.GetModifyPageType();
            var p = App.Current.ServiceProvider.GetRequiredService(t) as AbbreviationPage;

            p.SetModifyAbbreviation(abbrev);

            aaw.Content = p;
            aaw.Show();

            aaw.Closed += delegate { Init(); };
        }
    }
}
