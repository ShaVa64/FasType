using FasType.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace FasType.Models.Abbreviations
{
    [DebuggerDisplay("{" + nameof(ElementaryRepresentation) + "}")]
    public class SimpleAbbreviation : IAbbreviation
    {
        public string ShortForm { get; private set; }
        public string FullForm { get; private set; }

        public string ElementaryRepresentation => $"{ShortForm} -> {FullForm}";
        public string ComplexRepresentation => $"{ShortForm} -> {FullForm} \n{ShortForm.FirstCharToUpper()} -> {FullForm.FirstCharToUpper()}";

        public SimpleAbbreviation(string shortForm, string fullForm)
        {
            ShortForm = shortForm.ToLower();
            FullForm = fullForm.ToLower();
        }

        public virtual bool IsAbbreviation(string shortForm) => shortForm.ToLower() == ShortForm.ToLower();

        public virtual string GetFullForm(string shortForm)
        {
            if (shortForm.ToLower() == ShortForm.ToLower())
                return FullForm;
            return null;
        }

        public bool TryGetFullForm(string shortForm, out string fullForm)
        {
            fullForm = null;
            bool isAbrrev = IsAbbreviation(shortForm);
            if (!isAbrrev)
                return false;
            fullForm = GetFullForm(shortForm);
            return true;
        }
    }
}