using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Models.Abbreviations
{
    public class SimpleAbbreviation : BaseAbbreviation
    {
        public bool HasPlural => !string.IsNullOrEmpty(PluralForm);
        public bool HasGender => !string.IsNullOrEmpty(GenderForm);
        public bool HasGenderPlural => !string.IsNullOrEmpty(GenderPluralForm);

        public string GenderForm { get; private set; }
        public string PluralForm { get; private set; }
        public string GenderPluralForm { get; private set; }

        public SimpleAbbreviation(string shortForm, string fullForm, ulong used, string genderForm, string pluralForm, string genderPluralForm)
            : base(shortForm.ToLower(), fullForm, used)
        {
            PluralForm = pluralForm;
            GenderForm = genderForm;
            GenderPluralForm = genderPluralForm;
        }

        public override bool IsAbbreviation(string shortForm)
        {
            string sf = shortForm.ToLower();

            if (sf == ShortForm)
                return true;
            if (HasGender && sf == Linguistics.GenderType.Grammarify(ShortForm))
                return true;
            if (HasPlural && sf == Linguistics.PluralType.Grammarify(ShortForm))
                return true;
            if (HasGenderPlural && sf == Linguistics.GenderPluralType.Grammarify(ShortForm))
                return true;

            return false;
        }

        public override string? GetFullForm(string shortForm)
        {
            string sf = shortForm.ToLower();

            if (sf == ShortForm)
                return FullForm;
            if (HasGender && sf == Linguistics.GenderType.Grammarify(ShortForm))
                return GenderForm;
            if (HasPlural && sf == Linguistics.PluralType.Grammarify(ShortForm))
                return PluralForm;
            if (HasGenderPlural && sf == Linguistics.GenderPluralType.Grammarify(ShortForm))
                return GenderPluralForm;

            return null;
        }

        public override bool TryGetFullForm(string shortForm, [NotNullWhen(true)] out string? fullForm)
        {
            fullForm = null;
            bool isAbrrev = IsAbbreviation(shortForm);
            if (!isAbrrev)
                return false;
            fullForm = GetFullForm(shortForm) ?? throw new NullReferenceException();
            return true;
        }

        #region Representations
        static string ElementaryCapitalize(string @in)
        {
            StringBuilder sb = new();

            sb.Append('(')
                .Append(char.ToUpper(@in[0]))
                .Append('/')
                .Append(@in[0])
                .Append(')')
                .Append(@in[1..]);

            return sb.ToString();
        }

        protected override string GetElementaryRepresentation()
        {
            StringBuilder sb = new();

            string sf = ShortForm;
            string ff = FullForm;

            sb.Append(sf)
                .Append(SpacedArrow)
                .Append(ff);

            return sb.ToString();
        }

        protected override string GetComplexRepresentation()
        {
            StringBuilder sb = new();

            sb.Append(ShortForm)
                .Append(SpacedArrow)
                .Append(FullForm);
            if (HasGender)
            {
                sb.AppendLine()
                    .Append(Linguistics.GenderType.Grammarify(ShortForm))
                    .Append(SpacedArrow)
                    .Append(GenderForm);
            }

            if (HasPlural)
            {
                sb.AppendLine()
                    .Append(Linguistics.PluralType.Grammarify(ShortForm))
                    .Append(SpacedArrow)
                    .Append(PluralForm);
            }

            if (HasGenderPlural)
            {
                sb.AppendLine()
                    .Append(Linguistics.GenderPluralType.Grammarify(ShortForm))
                    .Append(SpacedArrow)
                    .Append(GenderPluralForm);
            }

            return sb.ToString();
        }
        #endregion
    }
}
