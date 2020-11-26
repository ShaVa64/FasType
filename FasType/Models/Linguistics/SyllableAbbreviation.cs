using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Models.Linguistics
{
    public record SyllableAbbreviation(string ShortForm, string FullForm, Position Position)
    {
        public bool IsBefore => Position.HasFlag(Position.Before);
        public bool IsIn => Position.HasFlag(Position.In);
        public bool IsAfter => Position.HasFlag(Position.After);
    }

    [Flags] public enum Position
    {
        Before = 1,
        In = 2,
        After = 4
    }
}
