using FasType.Core.Contexts;
using FasType.Core.Models.Dictionary;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Services
{
    public interface IDictionaryRepository : IGenericRepository<BaseDictionaryElement, string>
    {
    }

    public class DictionaryRepository : GenericRepository<BaseDictionaryElement, string>, IDictionaryRepository
    {
        public DictionaryRepository(DictionaryDbContext context) : base(context)
        { }
    }
}
