#nullable enable
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace FasType.Abbreviations
{

    [DebuggerDisplay("{" + nameof(ShortForm) + "} -> {" + nameof(FullForm) + "}")]
    public class SimpleAbbreviation : IAbbreviation
    {
        public string ShortForm { get; private set; }
        public string FullForm { get; private set; }

        public SimpleAbbreviation(string shortForm, string fullForm)
        {
            ShortForm = shortForm;
            FullForm = fullForm;
        }

        public virtual bool IsAbbreviation(string shortForm) => shortForm.ToLower() == ShortForm.ToLower();

        public virtual string? GetFullForm(string shortForm)
        {
            if (shortForm.ToLower() == ShortForm.ToLower())
                return FullForm;
            return null;
        }

        public bool TryGetFullForm(string shortForm, out string? fullForm)
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
