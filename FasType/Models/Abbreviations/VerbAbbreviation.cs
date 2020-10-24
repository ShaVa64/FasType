using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace FasType.Models.Abbreviations
{
    [DebuggerDisplay("{" + nameof(ElementaryRepresentation) + "}")]
    public class VerbAbbreviation : BaseAbbreviation
    {
        public VerbAbbreviation() 
            : base(null, null) { }

        public override string GetFullForm(string shortForm) => throw new NotImplementedException();
        public override bool IsAbbreviation(string shortForm) => throw new NotImplementedException();
        public override bool TryGetFullForm(string shortForm, out string fullForm) => throw new NotImplementedException();

        protected override string GetComplexRepresentation() => throw new NotImplementedException();
        protected override string GetElementaryRepresentation() => throw new NotImplementedException();
    }
}
