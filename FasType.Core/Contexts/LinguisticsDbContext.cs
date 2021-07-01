using FasType.Core.Models.Linguistics;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Contexts
{
    public class LinguisticsDbContext : DbContext
    {
        public DbSet<AbbreviationMethod> AbbreviationMethods { get; set; }
        public DbSet<GrammarType> GrammarTypes { get; set; }

        public LinguisticsDbContext(DbContextOptions<LinguisticsDbContext> options) : base(options)
        {
            _ = AbbreviationMethods ?? throw new NullReferenceException();
            _ = GrammarTypes ?? throw new NullReferenceException();
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<AbbreviationMethod>(b =>
            {
                b.HasKey(am => am.Key);

                b.Property(am => am.ShortForm).HasMaxLength(10).IsRequired();
                b.Property(am => am.FullForm).HasMaxLength(20).IsRequired();
                b.Property(am => am.Position);

                b.Ignore(am => am.IsBefore)
                 .Ignore(am => am.IsIn)
                 .Ignore(am => am.IsAfter);

                b.ToTable(nameof(AbbreviationMethods));
            });

            modelBuilder.Entity<GrammarType>(b =>
            {
                b.HasKey(gt => gt.Name);

                b.Property(gt => gt.Repr).HasMaxLength(4).IsRequired();

                b.ToTable(nameof(GrammarTypes));
            });
        }
    }
}
