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

        public SimpleAbbreviation(string shortForm, string fullForm)
            : base(shortForm.ToLower(), fullForm.ToLower()) { }

        public override bool IsAbbreviation(string shortForm) => shortForm.ToLower() == ShortForm.ToLower();

        public override string GetFullForm(string shortForm)
        {
            if (shortForm.ToLower() == ShortForm.ToLower())
                return FullForm;
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

        string ElementaryCapitalize(string @in)
        {
            StringBuilder sb = new();

            sb.Append("(");
            sb.Append(char.ToUpper(@in[0]));
            sb.Append("/");
            sb.Append(@in[0]);
            sb.Append(")");
            sb.Append(@in.Substring(1));

            return sb.ToString();
        }

        protected override string GetElementaryRepresentation()
        {
            StringBuilder sb = new();

            string sf = ElementaryCapitalize(ShortForm);
            string ff = ElementaryCapitalize(FullForm);
            
            sb.Append(sf);
            sb.Append(" -> ");
            sb.Append(ff);

            return sb.ToString();
        }

        protected override string GetComplexRepresentation()
        {
            StringBuilder sb = new();

            sb.Append(ShortForm);
            sb.Append(" -> ");
            sb.AppendLine(FullForm);

            sb.Append(ShortForm.FirstCharToUpper());
            sb.Append(" -> ");
            sb.Append(FullForm.FirstCharToUpper());

            return sb.ToString();
        }
    }
}