using FasType.Models.Abbreviations;
using FasType.Pages;
using FasType.Services;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Design;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Linq.Expressions;
using System.Text;

namespace FasType.Storage
{
    public class EFSqliteAbbreviationContext : DbContext, IAbbreviationStorage
    {
        protected static ILinguisticsStorage Linguistics => App.Current.ServiceProvider.GetRequiredService<ILinguisticsStorage>();
        public int Count => Abbreviations.Count();
        public DbSet<BaseAbbreviation> Abbreviations { get; set; }

        public Type ElementType => Abbreviations.AsNoTracking().ElementType;
        public Expression Expression => Abbreviations.AsNoTracking().Expression;
        public IQueryProvider Provider => Abbreviations.AsNoTracking().Provider;

        //public DbSet<VerbAbbreviation> VerbAbbreviations { get; set; }

        public EFSqliteAbbreviationContext(DbContextOptions<EFSqliteAbbreviationContext> options) : base(options) 
        {
            _ = Abbreviations ?? throw new NullReferenceException();
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<SimpleAbbreviation>();
            //modelBuilder.Entity<VerbAbbreviation>();
            modelBuilder.Entity<BaseAbbreviation>();
            //modelBuilder.Entity<VerbAbbreviation>().HasNoKey();
        }

        //public EFSqliteStorage(IConfiguration _configuration) : base()
        //{
        //    _filepath = _configuration.GetConnectionString("EFSqlite");
        //    OnConfiguring(new DbContextOptionsBuilder());
        //}
        //protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        //{
        //    optionsBuilder.UseSqlite(_filepath);
        //    base.OnConfiguring(optionsBuilder);
        //}

        public bool Add(BaseAbbreviation abbrev)
        {
            Abbreviations.Add(abbrev);
            //if (abbrev is SimpleAbbreviation simple)
            //    Abbreviations.Add(simple);
            //else if (abbrev is VerbAbbreviation verb)
            //    VerbAbbreviations.Add(verb);

            var r = SaveChanges();
            return r > 0;
            //return true;
        }

        public bool AddRange(IEnumerable<BaseAbbreviation> abbrevs)
        {
            Abbreviations.AddRange(abbrevs);

            var r = SaveChanges();
            return r > 0;
        }

        public bool Remove(BaseAbbreviation abbrev)
        {
            Abbreviations.Remove(abbrev);
            //if (abbrev is SimpleAbbreviation simple)
            //    SimpleAbbreviations.Remove(simple);
            //else if (abbrev is VerbAbbreviation verb)
            //    VerbAbbreviations.Remove(verb);

            var r = SaveChanges();
            return r > 0;
            //return true;
        }

        public bool Contains(BaseAbbreviation abbrev)
        {
            var b = Abbreviations.Where(a => a.FullForm == abbrev.FullForm /*&& a.ShortForm == abbrev.ShortForm*/).Count();
            return b > 0;
        }

        public bool Clear()
        {
            throw new NotImplementedException();

            //foreach (var abbrev in Abbreviations)
            //    Abbreviations.Remove(abbrev);
            //var r = SaveChanges();
            //return r > 0;
        }

        public bool UpdateUsed(BaseAbbreviation abbrev)
        {
            if (!Contains(abbrev))
                return false;
            abbrev.UpdateUsed();
            Abbreviations.Update(abbrev);

            var r = SaveChanges();
            return r > 0;
        }
        public bool UpdateAbbreviation(BaseAbbreviation abbrev)
        {
            if (!Abbreviations.Contains(abbrev) && Contains(abbrev))
            {
                var other = Abbreviations.Where(a => a.FullForm == abbrev.FullForm).SingleOrDefault();
                if (other == null)
                    return false;
                //if (other.Key != abbrev.Key)
                return Remove(other) && Add(abbrev);
                //var success = Remove(other);
                //if (!success)
                //    return false;
            }
            Abbreviations.Update(abbrev);

            var r = SaveChanges();
            return r > 0;
        }

        public IEnumerable<BaseAbbreviation> GetAbbreviations(string shortForm)
        {
            var forms = new List<string>() { shortForm };
            if (Linguistics.GenderType.TryUngrammarify(shortForm, out string? form))
                forms.Add(form);
            if (Linguistics.PluralType.TryUngrammarify(shortForm, out form))
                forms.Add(form);
            if (Linguistics.GenderPluralType.TryUngrammarify(shortForm, out form))
                forms.Add(form);

            var l = Abbreviations.Where(a => forms.Contains(a.ShortForm))/*.OrderByDescending(a => a.Used)*/.ToList();
            return l;
        }

        public IEnumerator<BaseAbbreviation> GetEnumerator() => Abbreviations.AsEnumerable().GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => Abbreviations.AsEnumerable().GetEnumerator();
    }

    class EFSqliteAbbreviationContextFactory : IDesignTimeDbContextFactory<EFSqliteAbbreviationContext>
    {
        public EFSqliteAbbreviationContext CreateDbContext(string[] args)
        {
            var optionsBuilder = new DbContextOptionsBuilder<EFSqliteAbbreviationContext>();
            optionsBuilder.UseSqlite("Data Source=abbreviation.db"); 

            return new EFSqliteAbbreviationContext(optionsBuilder.Options);
        }
    }
}