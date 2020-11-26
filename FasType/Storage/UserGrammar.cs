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
        public static GrammarTypeRecord GenderRecord { get; set; } = new(/*Properties.Resources.Gender,*/ "y", GrammarPosition.Postfix);
        public static GrammarTypeRecord PluralRecord { get; set; } = new(/*Properties.Resources.Plural,*/ "k", GrammarPosition.Postfix);
        public static GrammarTypeRecord GenderPluralRecord { get; set; } = new(/*Properties.Resources.Gender,*/ "w", GrammarPosition.Postfix);
    }
}
