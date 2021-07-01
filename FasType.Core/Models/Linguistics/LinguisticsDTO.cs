using FasType.Core.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Models.Linguistics
{
    public class LinguisticsDTO : ILinguisticsRepository
    {
        [NotNull] public GrammarType? GenderType { get; set; }
        [NotNull] public GrammarType? PluralType { get; set; }
        [NotNull] public GrammarType? GenderPluralType { get; set; }

        [NotNull] public IEnumerable<AbbreviationMethod>? AbbreviationMethods { get; set; }

        public int Count => throw new NotImplementedException();
        public void Add(AbbreviationMethod entity) => throw new NotImplementedException();
        public bool Contains(AbbreviationMethod entity) => throw new NotImplementedException();
        public bool Contains(Guid id) => throw new NotImplementedException();
        public void Dispose() => throw new NotImplementedException();
        public IEnumerable<AbbreviationMethod> GetAll() => throw new NotImplementedException();
        public AbbreviationMethod? GetById(Guid id) => throw new NotImplementedException();
        public void Remove(AbbreviationMethod entity) => throw new NotImplementedException();
        public void SaveChanges() => throw new NotImplementedException();
        public void Update(AbbreviationMethod entity) => throw new NotImplementedException();
        public IEnumerable<AbbreviationMethod> Where(Expression<Func<AbbreviationMethod, bool>> predicate) => throw new NotImplementedException();
        public string[] Words(string currentWord) => throw new NotImplementedException();
    }
}
