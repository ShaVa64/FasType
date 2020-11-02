using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace FasType.Models.Linguistics.Grammars
{
    public class GrammarType : ObservableObject
    {
        string _repr, _name;
        GrammarPosition _where;

        public GrammarPosition Where { get => _where; set => SetProperty(ref _where,value); }
        public string Repr { get => _repr; set => SetProperty(ref _repr, value); }
        public string Name { get => _name; set => SetProperty(ref _name, value); }

        public GrammarType(string name, string repr, GrammarPosition grammarPosition)
        {
            Name = name;
            Repr = repr;
            Where = grammarPosition;
        }

        public override string ToString() => $"{Name}: {(Where == GrammarPosition.Prefix ? $"{Repr}*" : $"*{Repr}")}";
    }

    public enum GrammarPosition
    {
        Prefix,
        Postfix
    }
}
