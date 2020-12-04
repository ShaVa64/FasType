using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Models.Linguistics
{
    public record AbbreviationMethodRecord(Guid Key, string ShortForm, string FullForm, SyllablePosition Position)
    {
        public bool IsBefore => Position.HasFlag(SyllablePosition.Before);
        public bool IsIn => Position.HasFlag(SyllablePosition.In);
        public bool IsAfter => Position.HasFlag(SyllablePosition.After);

        public override string ToString() => $"{ShortForm} {Utils.Unicodes.Arrow} {FullForm} ({Position})";
    }

    public class AbbreviationMethod : ObservableObject
    {
        string _shortForm, _fullForm;
        bool _isBefore, _isIn, _isAfter;

        public Guid Key { get; }
        public string ShortForm { get => _shortForm; set => SetProperty(ref _shortForm, value); }
        public string FullForm { get => _fullForm; set => SetProperty(ref _fullForm, value); }

        public bool IsBefore
        {
            get => _isBefore;
            set
            {
                SetProperty(ref _isBefore, value);
                OnPropertyChanged(nameof(Position));
            }
        }
        public bool IsIn
        {
            get => _isIn;
            set
            {
                SetProperty(ref _isIn, value);
                OnPropertyChanged(nameof(Position));
            }
        }
        public bool IsAfter
        {
            get => _isAfter;
            set
            {
                SetProperty(ref _isAfter, value);
                OnPropertyChanged(nameof(Position));
            }
        }

        public SyllablePosition Position => (SyllablePosition)((IsBefore ? 1 : 0) + (IsIn ? 2 : 0) + (IsAfter ? 4 : 0));

        public AbbreviationMethod(Guid key, string shortForm, string fullForm, SyllablePosition position)
        {
            Key = key;
            ShortForm = shortForm;
            FullForm = fullForm;
            IsBefore = position.HasFlag(SyllablePosition.Before);
            IsIn = position.HasFlag(SyllablePosition.In);
            IsAfter = position.HasFlag(SyllablePosition.After);
        }

        public override string ToString() => $"{ShortForm} {Utils.Unicodes.Arrow} {FullForm} ({Position})";

        public static explicit operator AbbreviationMethodRecord(AbbreviationMethod sa) => new(sa.Key, sa.ShortForm, sa.FullForm, sa.Position);
        public static explicit operator AbbreviationMethod(AbbreviationMethodRecord sar) => new(sar.Key, sar.ShortForm, sar.FullForm, sar.Position);
    }

    [Flags] public enum SyllablePosition
    {
        None = 0,
        Before = 1,
        In = 2,
        After = 4
    }
}
