using System;
using System.Collections.Generic;
using FasType.Models.Linguistics.Grammars;
using FasType.Services;

namespace FasType.Models.Linguistics
{
    public sealed class LinguisticsDTO : ILinguisticsStorage
    {
        public GrammarType GenderType { get; set; }
        public GrammarType PluralType { get; set; }
        public GrammarType GenderPluralType { get; set; }

        public GrammarType GenderCompletion { get; set; }
        public GrammarType PluralCompletion { get; set; }
        public GrammarType GenderPluralCompletion { get; set; }

        public IEnumerable<AbbreviationMethod> AbbreviationMethods { get; set; }

        public void Dispose() => throw new NotImplementedException();
        public bool Export(string filename) => throw new NotImplementedException();
        public bool Import(string filename) => throw new NotImplementedException();
    }
}
