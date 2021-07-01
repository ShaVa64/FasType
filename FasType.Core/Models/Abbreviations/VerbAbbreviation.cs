using FasType.Core.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Models.Abbreviations
{
    public class VerbAbbreviation : BaseAbbreviation
    {
        private VerbAbbreviation()
            : base(string.Empty, string.Empty, 0) { }

        public override string GetFullForm(string shortForm, ILinguisticsRepository linguistics) => throw new NotImplementedException();
        public override bool IsAbbreviation(string shortForm, ILinguisticsRepository linguistics) => throw new NotImplementedException();
        public override bool TryGetFullForm(string shortForm, ILinguisticsRepository linguistics, [NotNullWhen(true)] out string? fullForm) => throw new NotImplementedException();

        protected override string GetComplexRepresentation(ILinguisticsRepository linguistics) => throw new NotImplementedException();
        protected override string GetElementaryRepresentation() => throw new NotImplementedException();
    }
}
