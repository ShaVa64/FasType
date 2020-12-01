using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FasType.Models.Linguistics;
using FasType.Models.Linguistics.Grammars;

namespace FasType.Services
{
    public interface ILinguisticsStorage
    {
        //IEnumerable<GrammarType> GrammarTypes { get; set; }
        GrammarType GenderType { get; set; }
        GrammarType PluralType { get; set; }
        GrammarType GenderPluralType { get; set; }
        IEnumerable<SyllableAbbreviation> AbbreviationMethods { get; set; }


        bool Import(string filename);
        bool Export(string filename);
    }
}
