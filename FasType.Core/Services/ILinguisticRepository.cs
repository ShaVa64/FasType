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
    public interface ILinguisticRepository : IGenericRepository<AbbreviationMethod, Guid>
    {
    }

    public class LinguisticsRepository : GenericRepository<AbbreviationMethod, Guid>, ILinguisticRepository
    {
        public LinguisticsRepository(LinguisticsDbContext context) : base(context)
        {
        }
    }
}
