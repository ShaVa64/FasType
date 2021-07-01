using FasType.Models;
using FasType.Core.Models.Abbreviations;
using FasType.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using FasType.Core.Services;

namespace FasType.Pages
{
    /// <summary>
    /// Interaction logic for SimpleAbbreviationPage.xaml
    /// </summary>
    /// 

    public class AbbreviationPage : Page
    {
        public virtual void SetModifyAbbreviation(BaseAbbreviation ba)
        { }
        public virtual void SetNewAbbreviation(string shortForm, string fullForm, params string[] others)
        { }
    }

    public partial class SimpleAbbreviationPage : AbbreviationPage
    {
        private readonly IRepositoriesManager _repositories;
        private SimpleAbbreviationViewModel _currentVm;

        public SimpleAbbreviationPage(IRepositoriesManager repositories, AddSimpleAbbreviationViewModel addVm)
        {
            _repositories = repositories;
            InitializeComponent();

            DataContext = _currentVm = addVm;
            FirstTB.Focus();
            //Init();
        }

        public override void SetModifyAbbreviation(BaseAbbreviation ba) => SetModifyAbbreviation((SimpleAbbreviation)ba);
        public void SetModifyAbbreviation(SimpleAbbreviation abbrev) => DataContext = _currentVm = new ModifySimpleAbbreviationViewModel(_repositories, abbrev);
        public override void SetNewAbbreviation(string shortForm, string fullForm, params string[] others) => SetNewAbbreviation(shortForm, fullForm, others.Length >= 1 ? others[0] : string.Empty, others.Length >= 2 ? others[1] : string.Empty, others.Length >= 3 ? others[2] : string.Empty);
        public void SetNewAbbreviation(string shortForm, string fullForm, string genderForm, string pluralForm, string genderPluralForm)
        {
            DataContext = _currentVm = new AddSimpleAbbreviationViewModel(_repositories, shortForm, fullForm, genderForm, pluralForm, genderPluralForm);
            if (!string.IsNullOrEmpty(fullForm))
                MainButton.Focus();
        }
    }
}
