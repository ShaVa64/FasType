using FasType.Models.Abbreviations;
using FasType.Models.Dictionary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Services
{
    public interface IDictionaryStorage : IDisposable
    {
        int Count { get; }

        bool Contains(string fullForm);
        BaseDictionaryElement? GetElement(string fullForm);
        bool TryGetElement(string fullForm, out BaseDictionaryElement? s);

        BaseDictionaryElement[]? GetElements(string regexFullForm);
        bool TryGetElements(string fullForm, out BaseDictionaryElement[]? s);

        bool Add(BaseAbbreviation? abbrev);
        bool Add(BaseDictionaryElement? elem);

        T? GetElement<T>(string fullForm) where T : BaseDictionaryElement;
        bool TryGetElement<T>(string fullForm, out T? s) where T : BaseDictionaryElement;
    }
}
