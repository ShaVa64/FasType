using FasType.Models.Abbreviations;
using FasType.Pages;
using FasType.Services;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Design;
using Microsoft.Extensions.Configuration;
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
        public int Count => Abbreviations.Count();
        public DbSet<BaseAbbreviation> Abbreviations { get; set; }

        public Type ElementType => Abbreviations.AsQueryable().ElementType;
        public Expression Expression => Abbreviations.AsQueryable().Expression;
        public IQueryProvider Provider => Abbreviations.AsQueryable().Provider;

        //public DbSet<VerbAbbreviation> VerbAbbreviations { get; set; }

        public EFSqliteAbbreviationContext(DbContextOptions<EFSqliteAbbreviationContext> options) : base(options) { }

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
            var b = Abbreviations.Where(a => a.FullForm == abbrev.FullForm && a.ShortForm == abbrev.ShortForm).Count();
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

        public IEnumerable<BaseAbbreviation> GetAbbreviations(string shortForm)
        {
            var l = Abbreviations.AsEnumerable().Where(a =>  a.IsAbbreviation(shortForm)).OrderByDescending(a => a.Used).ToList();
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