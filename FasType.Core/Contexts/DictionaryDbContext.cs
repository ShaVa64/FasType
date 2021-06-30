using FasType.Core.Models.Dictionary;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Contexts
{
    public class DictionaryDbContext : DbContext
    {
        public DbSet<BaseDictionaryElement> Dictionary { get; set; }
        
        public DictionaryDbContext(DbContextOptions<DictionaryDbContext> options) : base(options)
        {
            _ = Dictionary ?? throw new NullReferenceException();
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<BaseDictionaryElement>(b =>
            {
                var splitStringConverter = new ValueConverter<string[], string>(v => string.Join(";", v), v => v.Split(';', StringSplitOptions.None));
                b.HasKey(e => e.FullForm);

                b.Property(e => e.FullForm).HasMaxLength(50).IsRequired();
                b.Property(e => e.Others).HasConversion(splitStringConverter, new Microsoft.EntityFrameworkCore.ChangeTracking.ArrayStructuralComparer<string>());
                b.Property(e => e.AllForms).HasConversion(splitStringConverter, new Microsoft.EntityFrameworkCore.ChangeTracking.ArrayStructuralComparer<string>());

                b.ToTable(nameof(Dictionary));
            });

            modelBuilder.Entity<SimpleDictionaryElement>(b =>
            {
                b.HasBaseType<BaseDictionaryElement>();

                b.Property(e => e.GenderForm).HasMaxLength(50);
                b.Property(e => e.PluralForm).HasMaxLength(50);
                b.Property(e => e.GenderPluralForm).HasMaxLength(50);
            });

            modelBuilder.Entity<VerbDictionaryElement>(b =>
            {
                b.HasBaseType<BaseDictionaryElement>();
            });
        }
    }
}
