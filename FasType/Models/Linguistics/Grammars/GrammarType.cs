using System;
using System.Collections.Generic;
using System.Text;

namespace FasType.Models.Linguistics.Grammars
{
    public class GrammarType
    {
        //public bool IsUsed { get; set; }
        public GrammarPosition Where { get; private set; }
        public string Repr { get; private set; }
        public string Name { get; private set; }

        public GrammarType(string name, string repr, GrammarPosition grammarPosition)
        {
            Name = name;
            Repr = repr;
            Where = grammarPosition;
        }
    }

    public enum GrammarPosition
    {
        Prefix,
        Postfix
    }
}
