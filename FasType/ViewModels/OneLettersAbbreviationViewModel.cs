using FasType.Core.Models;
using FasType.Core.Models.Abbreviations;
using FasType.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace FasType.ViewModels
{
    public class OneLettersAbbreviationViewModel : ObservableObject
    {
        private readonly IRepositoriesManager _repositories;

        public ICommand PageCommand { get; }
        public string ComplexRepresentation => Abbreviation.GetComplexRepresentation(_repositories.Linguistics);
        public string ButtonText { get; }
        public BaseAbbreviation Abbreviation { get; }

        public OneLettersAbbreviationViewModel(IRepositoriesManager repositories, BaseAbbreviation abbrev, ICommand pageCommand)
        {
            PageCommand = pageCommand;
            _repositories = repositories;
            Abbreviation = abbrev;

            ButtonText = string.IsNullOrEmpty(Abbreviation.FullForm) ? Properties.Resources.Add : Properties.Resources.Modify;
        }
    }
}
