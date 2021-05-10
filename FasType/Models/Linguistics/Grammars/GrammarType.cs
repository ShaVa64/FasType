using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Text;

namespace FasType.Models.Linguistics.Grammars
{
    public record GrammarTypeRecord(string Name, string Repr, GrammarPosition Position)
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

        public bool TryUngrammarify(string form, out string? shortForm) 
        {
            shortForm = null;
            if (!SuitsGrammar(form))
                return false;
            
            shortForm =  Position switch
            {
                GrammarPosition.Prefix => form[1..],
                GrammarPosition.Postfix => form[..^1],
                _ => throw new NotImplementedException()
            };
            return true;
        }

        public override string ToString() => Position == GrammarPosition.Prefix ? $"{Repr}*" : $"*{Repr}";
    }

    public class GrammarType : ObservableObject
    {
        string _repr;
        GrammarPosition _position;

        public string Name { get; }
        public string Repr { get => _repr; set => SetProperty(ref _repr, value); }
        public GrammarPosition Position { get => _position; set => SetProperty(ref _position,value); }

        public GrammarType(string name, string repr, GrammarPosition position)
        {
            Name = name;
            Repr = repr;
            _ = _repr ?? throw new NullReferenceException();
            Position = position;
        }

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
        public bool TryUngrammarify(string form, [NotNullWhen(true)][MaybeNullWhen(false)] out string? shortForm)
        {
            shortForm = null;
            if (!SuitsGrammar(form))
                return false;

            shortForm = Position switch
            {
                GrammarPosition.Prefix => form[(Repr.Length)..],
                GrammarPosition.Postfix => form[..^(Repr.Length)],
                _ => throw new NotImplementedException()
            };
            return true;
        }

        public override string ToString() => Position == GrammarPosition.Prefix ? $"{Repr}*" : $"*{Repr}";

        public static explicit operator GrammarTypeRecord(GrammarType gt) => new(gt.Name, gt.Repr, gt.Position);
        public static explicit operator GrammarType(GrammarTypeRecord gtr) => new(gtr.Name, gtr.Repr, gtr.Position);
    }

    public enum GrammarPosition
    {
        Prefix,
        Postfix
    }
}
