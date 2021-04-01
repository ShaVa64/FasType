using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Models.Dictionary
{
    [DebuggerDisplay("{" + nameof(DebuggerDisplay) + "}")]
    public abstract class BaseDictionaryElement
    {
        public static readonly BaseDictionaryElement OtherElement = new SimpleDictionaryElement(Properties.Resources.Other, "", "", "");
        public static readonly BaseDictionaryElement NoneElement = new SimpleDictionaryElement(Properties.Resources.None, "", "", "");

        [Key] public string FullForm { get; set; }
        public string[] Others { get; set; }
        public string[] AllForms { get; }
        string DebuggerDisplay => string.Join(", ", AllForms);

        protected BaseDictionaryElement(string fullForm, string[] others)
        {
            FullForm = fullForm;
            Others = others.ToArray();

            AllForms = Others.Prepend(fullForm).ToArray();
        }

        public override bool Equals(object? obj)
        {
            return obj is BaseDictionaryElement element &&
                   FullForm == element.FullForm;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(FullForm);
        }
    }

    public class SimpleDictionaryElement : BaseDictionaryElement
    {
        public string GenderForm { get; set; }
        public string PluralForm { get; set; }
        public string GenderPluralForm { get; set; }

        public SimpleDictionaryElement(string fullForm, string genderForm, string pluralForm, string genderPluralForm) : base(fullForm, new string[] { genderForm, pluralForm, genderPluralForm })
        {
            GenderForm = genderForm;
            PluralForm = pluralForm;
            GenderPluralForm = genderPluralForm;
        }

        public SimpleDictionaryElement(Abbreviations.SimpleAbbreviation sa) : this(sa.FullForm, sa.GenderForm, sa.PluralForm, sa.GenderPluralForm)
        { }
    }

    public class VerbDictionaryElement : BaseDictionaryElement
    {
        public VerbDictionaryElement(string fullform) : base(fullform, Array.Empty<string>())
        { }

        public VerbDictionaryElement(Abbreviations.VerbAbbreviation va) : this(va.FullForm)
        { }
    }
}
