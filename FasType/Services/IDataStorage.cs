using FasType.Models;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Services
{
    public interface IDataStorage
    {
        IEnumerable<IAbbreviation> GetAbbreviations(string shortForm);
        IAbbreviation GetAbbreviation(string shortForm);

        bool Add(IAbbreviation abbrev);
        Task<bool> AddAsync(IAbbreviation abbrev);
    }
}