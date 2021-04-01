﻿using FasType.Models.Abbreviations;
using FasType.Models.Dictionary;
using FasType.Services;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Design;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Storage
{
    public class EFSqliteDictionaryContext : DbContext, IDictionaryStorage
    {
        public DbSet<BaseDictionaryElement> Dictionary { get; set; }
        public int Count => Dictionary.Count();

        public EFSqliteDictionaryContext(DbContextOptions<EFSqliteDictionaryContext> options) : base(options) { }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            var splitStringConverter = new ValueConverter<string[], string>(v => string.Join(";", v), v => v.Split(new[] { ';' }));
            
            modelBuilder.Entity<SimpleDictionaryElement>();
            modelBuilder.Entity<BaseDictionaryElement>()
                .Property<string[]>(nameof(BaseDictionaryElement.Others))
                .HasConversion(splitStringConverter);
            modelBuilder.Entity<BaseDictionaryElement>()
                .Property(nameof(BaseDictionaryElement.AllForms))
                .HasConversion(splitStringConverter);
        }

        public bool Contains(string fullForm) => Dictionary.Find(fullForm) != null;
        public BaseDictionaryElement GetElement(string fullForm)
        {
            return Dictionary.Find(fullForm);
        }

        public bool TryGetElement(string fullForm, out BaseDictionaryElement s)
        {
            s = null;
            if (!Contains(fullForm))
                return false;
            s = GetElement(fullForm);
            return true;
        }

        public BaseDictionaryElement[] GetElements(string regexFullForm, int regexLength) => Dictionary.FromSqlRaw("SELECT * FROM Dictionary WHERE LOWER(FullForm) LIKE {0} AND LENGTH(FullForm) <= LENGTH({0}) + {1}", regexFullForm, regexLength).ToArray();
        public bool TryGetElements(string regexFullForm, out BaseDictionaryElement[] s)
        {
            s = null;
            var r = GetElements(regexFullForm, regexFullForm.Count(c => c == '%'));
            if (r == null || r.Length == 0)
                return false;
            s = r;
            return true;
        }

        public bool Add(BaseAbbreviation abbrev) => Add(abbrev switch
        {
            SimpleAbbreviation sa => new SimpleDictionaryElement(sa),
            VerbAbbreviation va => new VerbDictionaryElement(va),
            _ => null,
        });
        public bool Add(BaseDictionaryElement elem)
        {
            if (elem == null || elem.FullForm == Properties.Resources.Other)
                return false;
            Dictionary.Add(elem);

            int r = SaveChanges();
            return r > 0;
        }

        public T GetElement<T>(string fullForm) where T : BaseDictionaryElement => GetElement(fullForm) as T;
        public bool TryGetElement<T>(string fullForm, out T s) where T : BaseDictionaryElement => TryGetElement(fullForm, out s);
    }

    class EFSqliteDictionaryContextFactory : IDesignTimeDbContextFactory<EFSqliteDictionaryContext>
    {
        public EFSqliteDictionaryContext CreateDbContext(string[] args)
        {
            var optionsBuilder = new DbContextOptionsBuilder<EFSqliteDictionaryContext>();
            optionsBuilder.UseSqlite("Data Source=dictionary.db");

            return new EFSqliteDictionaryContext(optionsBuilder.Options);
        }
    }
}
