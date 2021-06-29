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
    }

    public class AbbreviationRepository : GenericRepository<BaseAbbreviation, Guid>, IAbbreviationRepository
    {
        public AbbreviationRepository(AbbreviationDbContext context) : base(context)
        { }
    }
}
