using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Models.Linguistics
{
    public class GrammarType : ObservableObject
    {
        string _repr;
        GrammarPosition _position;

        public string Name { get; }
        public string Repr { get => _repr; set => SetProperty(ref _repr, value); }
        public GrammarPosition Position { get => _position; set => SetProperty(ref _position, value); }

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
        public bool TryUngrammarify(string form, [NotNullWhen(true)] out string? shortForm)
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
    }

    public enum GrammarPosition
    {
        Prefix,
        Postfix
    }
}
