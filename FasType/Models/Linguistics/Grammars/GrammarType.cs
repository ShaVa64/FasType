using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace FasType.Models.Linguistics.Grammars
{
    public record GrammarTypeRecord(string Name, string Repr, GrammarPosition Position)
    {
        public override string ToString() => $"{Name}: {(Position == GrammarPosition.Prefix ? $"{Repr}*" : $"*{Repr}")}";
    }

    public class GrammarType : ObservableObject
    {
        string _repr, _name;
        GrammarPosition _position;

        public GrammarPosition Position { get => _position; set => SetProperty(ref _position,value); }
        public string Repr { get => _repr; set => SetProperty(ref _repr, value); }
        public string Name { get => _name; set => SetProperty(ref _name, value); }

        public GrammarType(string name, string repr, GrammarPosition grammarPosition)
        {
            Name = name;
            Repr = repr;
            Position = grammarPosition;
        }

        public override string ToString() => $"{Name}: {(Position == GrammarPosition.Prefix ? $"{Repr}*" : $"*{Repr}")}";
    
        public static explicit operator GrammarTypeRecord(GrammarType gt) => new(gt.Name, gt.Repr, gt.Position);
        public static explicit operator GrammarType(GrammarTypeRecord gtr) => new(gtr.Name, gtr.Repr, gtr.Position);
    }

    public enum GrammarPosition
    {
        Prefix,
        Postfix
    }
}
