using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using FasType.Models.Linguistics.Grammars;
using FasType.Services;

namespace FasType.Models.Linguistics
{
    public sealed class LinguisticsDTO : ILinguisticsStorage
    {
        [NotNull] public GrammarType? GenderType { get; set; }
        [NotNull] public GrammarType? PluralType { get; set; }
        [NotNull] public GrammarType? GenderPluralType { get; set; }

        [NotNull] public IEnumerable<AbbreviationMethod>? AbbreviationMethods { get; set; }

        public void Dispose() => throw new NotImplementedException();
        public bool Export(string filename) => throw new NotImplementedException();
        public bool Import(string filename) => throw new NotImplementedException();
        public string[] Words(string currentWord) => throw new NotImplementedException();
    }
}
