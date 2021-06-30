using FasType.Core.Contexts;
using FasType.Core.Models.Abbreviations;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Services
{
    public interface IAbbreviationRepository : IGenericRepository<BaseAbbreviation, Guid>
    {
        void UpdateUsed(BaseAbbreviation abbrev); 
        IEnumerable<BaseAbbreviation> this[string shortForm] => GetAbbreviations(shortForm);
        IEnumerable<BaseAbbreviation> GetAbbreviations(string shortForm);
    }

    public class AbbreviationRepository : GenericRepository<BaseAbbreviation, Guid, AbbreviationDbContext>, IAbbreviationRepository
    {
        readonly ILinguisticsRepository _linguistics;

        public AbbreviationRepository(ILinguisticsRepository linguistics, AbbreviationDbContext context) : base(context)
        {
            _linguistics = linguistics;
        }

        public void UpdateUsed(BaseAbbreviation abbrev)
        {
            abbrev.UpdateUsed();
            Update(abbrev);
        }

        public IEnumerable<BaseAbbreviation> GetAbbreviations(string shortForm)
        {
            var forms = new List<string>() { shortForm };
            if (_linguistics.GenderType.TryUngrammarify(shortForm, out string? form))
                forms.Add(form);
            if (_linguistics.PluralType.TryUngrammarify(shortForm, out form))
                forms.Add(form);
            if (_linguistics.GenderPluralType.TryUngrammarify(shortForm, out form))
                forms.Add(form);

            var l = Where(a => forms.Contains(a.ShortForm))/*.OrderByDescending(a => a.Used)*/.ToList().Where(ba => ba.IsAbbreviation(shortForm, _linguistics));
            return l;
        }
    }
}
