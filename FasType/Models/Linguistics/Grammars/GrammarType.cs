using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace FasType.Models.Linguistics.Grammars
{
    public record GrammarTypeRecord(string Repr, GrammarPosition Position)
    {
        public string Grammarify(string form) => Position switch
        {
            GrammarPosition.Prefix => Repr + form,
            GrammarPosition.Postfix => form + Repr,
            _ => throw new NotImplementedException()
        };

        public bool SuitsGrammar(string form) => Position switch
        {
            GrammarPosition.Prefix => form.StartsWith(Repr),
            GrammarPosition.Postfix => form.EndsWith(Repr),
            _ => throw new NotImplementedException()
        };

        public override string ToString() => Position == GrammarPosition.Prefix ? $"{Repr}*" : $"*{Repr}";
    }

    public class GrammarType : ObservableObject
    {
        string _repr;
        GrammarPosition _position;

        public string Repr { get => _repr; set => SetProperty(ref _repr, value); }
        public GrammarPosition Position { get => _position; set => SetProperty(ref _position,value); }

        public GrammarType(string repr, GrammarPosition grammarPosition)
        {
            Repr = repr;
            Position = grammarPosition;
        }

        public override string ToString() => Position == GrammarPosition.Prefix ? $"{Repr}*" : $"*{Repr}";

        public static explicit operator GrammarTypeRecord(GrammarType gt) => new(gt.Repr, gt.Position);
        public static explicit operator GrammarType(GrammarTypeRecord gtr) => new(gtr.Repr, gtr.Position);
    }

    public enum GrammarPosition
    {
        Prefix,
        Postfix
    }
}
