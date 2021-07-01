using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Services
{
    public interface IGenericRepository<T, TId>
    {
        int Count { get; }

        void Add(T entity);
        bool Contains(T entity);
        bool Contains(TId id);
        IEnumerable<T> Where(Expression<Func<T, bool>> predicate);
        T? GetById(TId id);
        IEnumerable<T> GetAll();
        void Remove(T entity);
        void Update(T entity);
        void SaveChanges();
    }

    public abstract class GenericRepository<T, TId, TContext> : IGenericRepository<T, TId>
        where T : class
        where TContext : DbContext
    {
        protected readonly TContext _context;

        protected DbSet<T> Set => _context.Set<T>();
        public int Count => Set.Count();

        public GenericRepository(TContext context)
        {
            _context = context;
        }

        public virtual void Add(T entity) => Set.Add(entity);
        public virtual bool Contains(T entity) => Set.Contains(entity);
        public virtual bool Contains(TId id) => Set.Find(id) != null;
        public virtual IEnumerable<T> Where(Expression<Func<T, bool>> predicate) => Set.Where(predicate);
        public virtual T? GetById(TId id) => Set.Find(id);
        public IEnumerable<T> GetAll() => Set.ToArray();
        public virtual void Remove(T entity) => Set.Remove(entity);
        public virtual void Update(T entity) => Set.Update(entity);
        public virtual void SaveChanges() => _context.SaveChanges();
    }
}
