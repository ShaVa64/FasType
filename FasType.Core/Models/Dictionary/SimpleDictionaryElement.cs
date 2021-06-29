using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Models.Dictionary
{
    public class SimpleDictionaryElement : BaseDictionaryElement
    {
        public string GenderForm { get; set; }
        public string PluralForm { get; set; }
        public string GenderPluralForm { get; set; }

        public SimpleDictionaryElement(string fullForm, string genderForm, string pluralForm, string genderPluralForm) : base(fullForm, new string[] { genderForm, pluralForm, genderPluralForm })
        {
            GenderForm = genderForm;
            PluralForm = pluralForm;
            GenderPluralForm = genderPluralForm;
        }

        public SimpleDictionaryElement(Abbreviations.SimpleAbbreviation sa) : this(sa.FullForm, sa.GenderForm, sa.PluralForm, sa.GenderPluralForm)
        { }
    }
}
