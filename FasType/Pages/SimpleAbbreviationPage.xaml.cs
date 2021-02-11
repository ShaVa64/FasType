using FasType.Models;
using FasType.Models.Abbreviations;
using FasType.Services;
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
        SimpleAbbreviationViewModel _currentVm;

        public SimpleAbbreviationPage(AddSimpleAbbreviationViewModel addVm)
        {
            InitializeComponent();

            DataContext = _currentVm = addVm;
            FirstTB.Focus();

            //Init();
        }

        public override void SetModifyAbbreviation(BaseAbbreviation ba) => SetModifyAbbreviation(ba as SimpleAbbreviation);
        public void SetModifyAbbreviation(SimpleAbbreviation abbrev)
        {
            DataContext = _currentVm = new ModifySimpleAbbreviationViewModel(abbrev);

            //_currentVm.ShortForm = abbrev.ShortForm;
            //_currentVm.FullForm = abbrev.FullForm;
            //_currentVm.GenderForm = abbrev.GenderForm;
            //_currentVm.PluralForm = abbrev.PluralForm;
            //_currentVm.GenderPluralForm = abbrev.GenderPluralForm;
        }
        public override void SetNewAbbreviation(string shortForm, string fullForm, params string[] others) => SetNewAbbreviation(shortForm, fullForm, others[0], others[1], others[2]);
        public void SetNewAbbreviation(string shortForm, string fullForm, string genderForm, string pluralForm, string genderPluralForm)
        {
            DataContext = _currentVm = new AddSimpleAbbreviationViewModel(shortForm, fullForm, genderForm, pluralForm, genderPluralForm);
        }
    }
}
