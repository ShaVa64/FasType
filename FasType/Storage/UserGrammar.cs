using FasType.Models.Linguistics;
using FasType.Models.Linguistics.Grammars;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Storage
{
    public static class UserGrammar
    {
        public static GrammarTypeRecord GenderRecord { get; set; } = new("Gender", "y", GrammarPosition.Postfix);
        public static GrammarTypeRecord PluralRecord { get; set; } = new("Plural", "k", GrammarPosition.Postfix);
        public static GrammarTypeRecord GenderPluralRecord { get; set; } = new("GenderPlural", "w", GrammarPosition.Postfix);

        public static SyllableAbbreviationRecord[] SyllabesAbbreviations { get; set; } = 
        { 
            new(Guid.NewGuid(), "è", "an", SyllablePosition.Before | SyllablePosition.In | SyllablePosition.After), 
            new(Guid.NewGuid(), "o", "au", SyllablePosition.Before | SyllablePosition.In) 
        };
    }
}
