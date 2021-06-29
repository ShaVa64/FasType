using FasType.Core.Contexts;
using FasType.Core.Models.Linguistics;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Services
{
    public interface ILinguisticsRepository : IGenericRepository<AbbreviationMethod, Guid>
    {
        public GrammarType GenderType { get; set; }
        public GrammarType PluralType { get; set; }
        public GrammarType GenderPluralType { get; set; }
    }

    public class LinguisticsRepository : GenericRepository<AbbreviationMethod, Guid>, ILinguisticsRepository
    {
        public GrammarType GenderType { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public GrammarType PluralType { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public GrammarType GenderPluralType { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        
        public LinguisticsRepository(LinguisticsDbContext context) : base(context)
        {
        }
    }
}
