using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using FasType.Services;
using System.Linq.Expressions;
using FasType.Models.Linguistics;
using FasType.Models.Linguistics.Grammars;
using System.Collections;
using Microsoft.EntityFrameworkCore.Design;
using FasType.Utils;
using System.Linq;
using System.Runtime.CompilerServices;

namespace FasType.Storage
{
    public class EFSqliteLinguisticsContext : DbContext, ILinguisticsStorage
    {
        public DbSet<GrammarTypeRecord> GrammarTypes { get; set; }
        public DbSet<SyllableAbbreviationRecord> AbbreviationMethods { get; set; }

        IEnumerable<SyllableAbbreviation> ILinguisticsStorage.AbbreviationMethods 
        { 
            get => GetAbbreviationMethods();
            set => SetAbbreviationMethods(value.Cast<SyllableAbbreviationRecord>());
        }
        public GrammarType GenderType { get => GetGrammarType(); set => SetGrammarType((GrammarTypeRecord)value); }
        public GrammarType PluralType { get => GetGrammarType(); set => SetGrammarType((GrammarTypeRecord)value); }
        public GrammarType GenderPluralType { get => GetGrammarType(); set => SetGrammarType((GrammarTypeRecord)value); }

        public EFSqliteLinguisticsContext(DbContextOptions<EFSqliteLinguisticsContext> options) : base(options) { }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<GrammarTypeRecord>().HasKey(nameof(GrammarTypeRecord.Name));
            modelBuilder.Entity<SyllableAbbreviationRecord>().HasKey(nameof(SyllableAbbreviationRecord.Key));

            base.OnModelCreating(modelBuilder);
        }

        IEnumerable<SyllableAbbreviation> GetAbbreviationMethods()
        {
            var enumerable = AbbreviationMethods.AsEnumerable().Cast<SyllableAbbreviation>();

            return enumerable;
        }
        void SetAbbreviationMethods(IEnumerable<SyllableAbbreviationRecord> enumerable)
        {
            AbbreviationMethods.RemoveRange(AbbreviationMethods);
            AbbreviationMethods.AddRange(enumerable);

            SaveChanges();
        }
        GrammarType GetGrammarType([CallerMemberName] string name = null)
        {
            var gtr = GrammarTypes.Find(name);

            return (GrammarType)(gtr ?? new(name, "", GrammarPosition.Prefix));
        }
        void SetGrammarType(GrammarTypeRecord gtr)
        {
            GrammarTypeRecord record = GrammarTypes.Find(gtr.Name);
            if (record is not null)
                GrammarTypes.Remove(record);
            GrammarTypes.Add(gtr);

            SaveChanges();
        }

        public bool Import(string filename) => throw new NotImplementedException();
        public bool Export(string filename) => throw new NotImplementedException();
    }

    class EFSqliteLinguisticsContextFactory : IDesignTimeDbContextFactory<EFSqliteLinguisticsContext>
    {
        public EFSqliteLinguisticsContext CreateDbContext(string[] args)
        {
            var optionsBuilder = new DbContextOptionsBuilder<EFSqliteLinguisticsContext>();
            optionsBuilder.UseSqlite("Data Source=linguistics.db");

            return new EFSqliteLinguisticsContext(optionsBuilder.Options);
        }
    }
}
