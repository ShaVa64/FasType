using FasType.Models.Abbreviations;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Services
{
    public interface IDataStorage : IEnumerable<BaseAbbreviation>
    {
        int Count { get; }

        IEnumerable<BaseAbbreviation> this[string shortForm] => GetAbbreviations(shortForm);
        IEnumerable<BaseAbbreviation> GetAbbreviations(string shortForm);

        bool Add(BaseAbbreviation abbrev);
        //Task<bool> AddAsync(IAbbreviation abbrev);

        bool Remove(BaseAbbreviation abbrev);
        bool Clear();
    }
}