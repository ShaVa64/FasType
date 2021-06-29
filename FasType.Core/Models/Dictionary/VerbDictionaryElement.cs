using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Models.Dictionary
{
    public class VerbDictionaryElement : BaseDictionaryElement
    {
        public VerbDictionaryElement(string fullform) : base(fullform, Array.Empty<string>())
        { }

        public VerbDictionaryElement(Abbreviations.VerbAbbreviation va) : this(va.FullForm)
        { }
    }
}
