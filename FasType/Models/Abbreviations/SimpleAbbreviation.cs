using FasType.Storage;
using FasType.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace FasType.Models.Abbreviations
{
    public class SimpleAbbreviation : BaseAbbreviation
    {

        //string IAbbreviation.ShortForm => ShortForm;
        //string IAbbreviation.FullForm => FullForm;

        //public string ElementaryRepresentation => GetElementaryRepresentation();//$"{ShortForm} -> {FullForm}";
        //public string ComplexRepresentation => GetElementaryRepresentation();// GetComplexRepresentation();//$"{ShortForm} -> {FullForm}{Environment.NewLine}{ShortForm.FirstCharToUpper()} -> {FullForm.FirstCharToUpper()}";

        public bool HasPlural => !string.IsNullOrEmpty(PluralForm);
        public bool HasGender => !string.IsNullOrEmpty(GenderForm);
        public bool HasGenderPlural => !string.IsNullOrEmpty(GenderPluralForm);

        [Required] [Column(TypeName = "varchar(50)")] public string GenderForm { get; private set; }
        [Required] [Column(TypeName = "varchar(50)")] public string PluralForm { get; private set; }
        [Required] [Column(TypeName = "varchar(50)")] public string GenderPluralForm { get; private set; }

        public SimpleAbbreviation(string shortForm, string fullForm, ulong used, string genderForm, string pluralForm, string genderPluralForm)
            : base(shortForm.ToLower(), fullForm.ToLower(), used) 
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

        public override string GetFullForm(string shortForm)
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

        public override bool TryGetFullForm(string shortForm, out string fullForm)
        {
            fullForm = null;
            bool isAbrrev = IsAbbreviation(shortForm);
            if (!isAbrrev)
                return false;
            fullForm = GetFullForm(shortForm);
            return true;
        }

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

            string sf = ShortForm;//ElementaryCapitalize(ShortForm);
            string ff = FullForm ;//ElementaryCapitalize(FullForm);
            
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

            //sb.AppendLine()
            //    .Append(ShortForm.FirstCharToUpper())
            //    .Append(SpacedArrow)
            //    .Append(FullForm.FirstCharToUpper());
            
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
    }
}