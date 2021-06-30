using FasType.Core.Models.Abbreviations;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Contexts
{
    public class AbbreviationDbContext : DbContext
    {
        public DbSet<BaseAbbreviation> Abbreviations { get; set; }

        public AbbreviationDbContext(DbContextOptions<AbbreviationDbContext> options) : base(options)
        {
            _ = Abbreviations ?? throw new NullReferenceException();
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<BaseAbbreviation>(b =>
            {
                b.HasKey(a => a.Id);

                b.Property(a => a.Id).ValueGeneratedOnAdd().IsRequired();
                b.Property(a => a.ShortForm).HasMaxLength(50).IsRequired();
                b.Property(a => a.FullForm).HasMaxLength(50).IsRequired();
                b.Property(a => a.Used).IsRequired();

                b.ToTable(nameof(Abbreviations));
            });

            modelBuilder.Entity<SimpleAbbreviation>(b =>
            {
                b.HasBaseType<BaseAbbreviation>();

                b.Property(a => a.GenderForm).HasMaxLength(50);
                b.Property(a => a.PluralForm).HasMaxLength(50);
                b.Property(a => a.GenderPluralForm).HasMaxLength(50);
            });

            modelBuilder.Entity<VerbAbbreviation>(b =>
            {
                b.HasBaseType<BaseAbbreviation>();
            });
        }
    }
}
