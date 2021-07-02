using FasType.Core.Services;
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

        public override bool IsAbbreviation(string shortForm, ILinguisticsRepository linguistics)
        {
            string sf = shortForm.ToLower();

            if (sf == ShortForm)
                return true;
            else if (HasGender && sf == linguistics.GenderType.Grammarify(ShortForm))
                return true;
            else if (HasPlural && sf == linguistics.PluralType.Grammarify(ShortForm))
                return true;
            else if (HasGenderPlural && sf == linguistics.GenderPluralType.Grammarify(ShortForm))
                return true;

            return false;
        }

        public override string? GetFullForm(string shortForm, ILinguisticsRepository linguistics)
        {
            string sf = shortForm.ToLower();

            if (sf == ShortForm)
                return FullForm;
            else if (HasGender && sf == linguistics.GenderType.Grammarify(ShortForm))
                return GenderForm;
            else if (HasPlural && sf == linguistics.PluralType.Grammarify(ShortForm))
                return PluralForm;
            else if (HasGenderPlural && sf == linguistics.GenderPluralType.Grammarify(ShortForm))
                return GenderPluralForm;

            return null;
        }

        public override bool TryGetFullForm(string shortForm, ILinguisticsRepository linguistics, [NotNullWhen(true)] out string? fullForm)
        {
            fullForm = null;
            bool isAbrrev = IsAbbreviation(shortForm, linguistics);
            if (!isAbrrev)
                return false;
            fullForm = GetFullForm(shortForm, linguistics) ?? throw new NullReferenceException();
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

        public override string GetComplexRepresentation(ILinguisticsRepository linguistics)
        {
            StringBuilder sb = new();

            sb.Append(ShortForm)
                .Append(SpacedArrow)
                .Append(FullForm);
            if (HasGender)
            {
                sb.AppendLine()
                    .Append(linguistics.GenderType.Grammarify(ShortForm))
                    .Append(SpacedArrow)
                    .Append(GenderForm);
            }

            if (HasPlural)
            {
                sb.AppendLine()
                    .Append(linguistics.PluralType.Grammarify(ShortForm))
                    .Append(SpacedArrow)
                    .Append(PluralForm);
            }

            if (HasGenderPlural)
            {
                sb.AppendLine()
                    .Append(linguistics.GenderPluralType.Grammarify(ShortForm))
                    .Append(SpacedArrow)
                    .Append(GenderPluralForm);
            }

            return sb.ToString();
        }
        #endregion
    }
}
