using FasType.Models;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Services
{
    public interface IDataStorage : ICollection<IAbbreviation>
    {
        bool ICollection<IAbbreviation>.IsReadOnly => false;

        IEnumerable<IAbbreviation> this[string shortForm] => GetAbbreviations(shortForm);

        IEnumerable<IAbbreviation> GetAbbreviations(string shortForm);
        IAbbreviation GetAbbreviation(string shortForm);

        new bool Add(IAbbreviation abbrev);
        Task<bool> AddAsync(IAbbreviation abbrev);
    }
}