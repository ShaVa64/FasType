using FasType.Models.Abbreviations;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Services
{
    public interface IDataStorage : IEnumerable<IAbbreviation>
    {
        int Count { get; }

        IEnumerable<IAbbreviation> this[string shortForm] => GetAbbreviations(shortForm);
        IEnumerable<IAbbreviation> GetAbbreviations(string shortForm);

        bool Add(IAbbreviation abbrev);
        //Task<bool> AddAsync(IAbbreviation abbrev);

        bool Remove(IAbbreviation abbrev);
        bool Clear();
    }
}