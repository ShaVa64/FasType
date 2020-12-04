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
using Microsoft.Extensions.DependencyInjection;
using FasType.Models.Abbreviations;

namespace FasType.Storage
{
    public class EFSqliteLinguisticsContext : DbContext, ILinguisticsStorage
    {
        public DbSet<GrammarTypeRecord> GrammarTypes { get; set; }
        public DbSet<AbbreviationMethodRecord> AbbreviationMethods { get; set; }

        IEnumerable<AbbreviationMethod> ILinguisticsStorage.AbbreviationMethods 
        { 
            get => GetAbbreviationMethods();
            set => SetAbbreviationMethods(value.Cast<AbbreviationMethodRecord>());
        }
        public GrammarType GenderType { get => GetGrammarType(); set => SetGrammarType((GrammarTypeRecord)value); }
        public GrammarType PluralType { get => GetGrammarType(); set => SetGrammarType((GrammarTypeRecord)value); }
        public GrammarType GenderPluralType { get => GetGrammarType(); set => SetGrammarType((GrammarTypeRecord)value); }

        public GrammarType GenderCompletion { get => GetGrammarType(); set => SetGrammarType((GrammarTypeRecord)value); }
        public GrammarType PluralCompletion { get => GetGrammarType(); set => SetGrammarType((GrammarTypeRecord)value); }
        public GrammarType GenderPluralCompletion { get => GetGrammarType(); set => SetGrammarType((GrammarTypeRecord)value); }

        public EFSqliteLinguisticsContext(DbContextOptions<EFSqliteLinguisticsContext> options) : base(options) { }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<GrammarTypeRecord>().HasKey(nameof(GrammarTypeRecord.Name));
            modelBuilder.Entity<AbbreviationMethodRecord>().HasKey(nameof(AbbreviationMethodRecord.Key));

            base.OnModelCreating(modelBuilder);
        }

        IEnumerable<AbbreviationMethod> GetAbbreviationMethods()
        {
            var enumerable = AbbreviationMethods.AsEnumerable().Cast<AbbreviationMethod>();

            return enumerable;
        }
        void SetAbbreviationMethods(IEnumerable<AbbreviationMethodRecord> enumerable)
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
            if (gtr == record)
                return;
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
