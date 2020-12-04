using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FasType.Models.Linguistics;
using FasType.Models.Linguistics.Grammars;

namespace FasType.Services
{
    public interface ILinguisticsStorage : IDisposable
    {
        //IEnumerable<GrammarType> GrammarTypes { get; set; }
        GrammarType GenderType { get; set; }
        GrammarType PluralType { get; set; }
        GrammarType GenderPluralType { get; set; }

        GrammarType GenderCompletion { get; set; }
        GrammarType PluralCompletion { get; set; }
        GrammarType GenderPluralCompletion { get; set; }

        IEnumerable<AbbreviationMethod> AbbreviationMethods { get; set; }


        bool Import(string filename);
        bool Export(string filename);
    }
}
