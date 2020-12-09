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

        List<string> Words(string _curr, string from, List<string> poss)
        {
            if (from == string.Empty)
            {
                poss.Add(_curr);
                return poss;
            }

            AbbreviationMethodRecord[] amrs = Array.Empty<AbbreviationMethodRecord>();
            if (_curr == string.Empty)
            {
                amrs = AbbreviationMethods.Where(m => m.Position.HasFlag(SyllablePosition.Before) && from.StartsWith(m.ShortForm)).ToArray();//.AsEnumerable().Where(m => m.SatisfiesBefore(from)).ToArray();
            }
            else
            {
                amrs = AbbreviationMethods.Where(m => m.Position.HasFlag(SyllablePosition.In) && from.StartsWith(m.ShortForm) && from.Length != m.ShortForm.Length).ToArray();
                amrs = amrs.Concat(AbbreviationMethods.Where(m => m.Position.HasFlag(SyllablePosition.After) && from.EndsWith(m.ShortForm) && from.Length == m.ShortForm.Length)).ToArray();
            }

            string temp = _curr;
            _curr += from[0];
            Words(_curr, from[1..], poss);

            _curr = temp;
            foreach (var amr in amrs)
                Words(_curr + amr.FullForm, from[(amr.ShortForm.Length)..], poss);

            return poss;
        }

        public List<string> Words(string currentWord) => Words("", currentWord, new());

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
