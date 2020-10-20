using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace FasType.Models.Abbreviations
{
    [DebuggerDisplay("{" + nameof(ElementaryRepresentation) + "}")]
    public class VerbAbbreviation : IAbbreviation
    {
        public string FullForm => throw new NotImplementedException();
        public string ShortForm => throw new NotImplementedException();
        public string ElementaryRepresentation => throw new NotImplementedException();
        public string ComplexRepresentation => throw new NotImplementedException();

        public string GetFullForm(string shortForm) => throw new NotImplementedException();
        public bool IsAbbreviation(string shortForm) => throw new NotImplementedException();
        public bool TryGetFullForm(string shortForm, out string fullForm) => throw new NotImplementedException();
    }
}
