using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Models.Abbreviations
{
    public class VerbAbbreviation : BaseAbbreviation
    {
        private VerbAbbreviation()
            : base(string.Empty, string.Empty, 0) { }

        public override string GetFullForm(string shortForm) => throw new NotImplementedException();
        public override bool IsAbbreviation(string shortForm) => throw new NotImplementedException();
        public override bool TryGetFullForm(string shortForm, out string fullForm) => throw new NotImplementedException();

        protected override string GetComplexRepresentation() => throw new NotImplementedException();
        protected override string GetElementaryRepresentation() => throw new NotImplementedException();
    }
}
