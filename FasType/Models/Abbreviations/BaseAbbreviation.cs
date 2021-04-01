using FasType.Services;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;

namespace FasType.Models.Abbreviations
{
    [Table("Abbreviations")]
    [DebuggerDisplay("{" + nameof(ElementaryRepresentation) + "}")]
    public abstract class BaseAbbreviation
    {
        public static readonly BaseAbbreviation OtherAbbreviation = new SimpleAbbreviation("", Properties.Resources.Other, 0, "", "", "");

        protected static ILinguisticsStorage Linguistics => App.Current.ServiceProvider.GetRequiredService<ILinguisticsStorage>();

        protected readonly static int _stringKeyLength = 2;
        protected readonly static string SpacedArrow = $" {Utils.Unicodes.Arrow} ";

        [Key] public Guid Key { get; private set; }
        [Required, MaxLength(50)] public string ShortForm { get; private set; }
        [Required, MaxLength(50)] public string FullForm { get; private set; }
        [Required] public ulong Used { get; private set; }

        public string StringKey => string.Concat(ShortForm.Take(_stringKeyLength));

        public BaseAbbreviation(string shortForm, string fullForm, ulong used)
        {
            Key = Guid.NewGuid();
            ShortForm = shortForm;
            FullForm = fullForm;
            Used = used;
        }

        public string ElementaryRepresentation => GetElementaryRepresentation();
        public string ComplexRepresentation => GetComplexRepresentation();

        public void UpdateUsed() => Used++;

        public abstract bool IsAbbreviation(string shortForm);
        public abstract string? GetFullForm(string shortForm);
        public abstract bool TryGetFullForm(string shortForm, [MaybeNullWhen(false)] out string? fullForm);

        protected abstract string GetElementaryRepresentation();
        protected abstract string GetComplexRepresentation();
    }
}
