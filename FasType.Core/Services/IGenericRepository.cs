using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Services
{
    public interface IGenericRepository<T, U>
    {
        int Count { get; }

        void Add(T entity);
        bool Contains(T entity);
        IEnumerable<T> Where(Expression<Func<T, bool>> predicate);
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
        public int Count => Set.Count();

        public GenericRepository(DbContext context)
        {
            _context = context;
        }

        public virtual void Add(T entity) => Set.Add(entity);
        public virtual bool Contains(T entity) => Set.Contains(entity);
        public virtual IEnumerable<T> Where(Expression<Func<T, bool>> predicate) => Set.Where(predicate);
        public virtual T GetById(U id) => Set.Find(id);
        public IEnumerable<T> GetAll() => Set.ToArray();
        public virtual void Remove(T entity) => Set.Remove(entity);
        public virtual void Update(T entity) => Set.Update(entity);
        public virtual void SaveChanges() => _context.SaveChanges();

    }
}
