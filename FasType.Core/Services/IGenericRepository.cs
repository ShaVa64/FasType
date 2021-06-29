using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Services
{
    public interface IGenericRepository<T, U>
    {
        void Add(T entity);
        T GetById(U id);
        IEnumerable<T> GetAll();
        void Remove(T entity);
        void Update(T entity);
        void SaveChanges();
    }

    public abstract class GenericRepository<T, U> : IGenericRepository<T, U>
        where T : class
    {
        private readonly DbContext _context;

        protected DbSet<T> Set => _context.Set<T>();
        
        public GenericRepository(DbContext context)
        {
            _context = context;
        }

        public virtual void Add(T entity) => Set.Add(entity);
        public virtual T GetById(U id) => Set.Find(id);
        public IEnumerable<T> GetAll() => Set.ToArray();
        public virtual void Remove(T entity) => Set.Remove(entity);
        public virtual void Update(T entity) => Set.Update(entity);
        public virtual void SaveChanges() => _context.SaveChanges();

    }
}
