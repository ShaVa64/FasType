using FasType.Models.Abbreviations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Services
{
    public interface IAbbreviationStorage : IQueryable<BaseAbbreviation>, IEnumerable<BaseAbbreviation>, IDisposable
    {
        int Count { get; }

        IEnumerable<BaseAbbreviation> this[string shortForm] => GetAbbreviations(shortForm);
        IEnumerable<BaseAbbreviation> GetAbbreviations(string shortForm);

        bool UpdateAbbreviation(BaseAbbreviation abbrev);
        Task<bool> UpdateUsedAsync(BaseAbbreviation abbrev);
        bool Add(BaseAbbreviation abbrev);
        bool AddRange(IEnumerable<BaseAbbreviation> abbrevs);
        bool Contains(BaseAbbreviation abbrev);
        bool Remove(BaseAbbreviation abbrev);
        bool Clear();
    }
}