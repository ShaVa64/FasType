using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;

namespace FasType.Models.Abbreviations
{
    [Table("Abbreviations")]
    public abstract class BaseAbbreviation
    {
        [Key] public Guid Key { get; private set; }
        [Required] [Column(TypeName = "varchar(50)")] public string ShortForm { get; private set; }
        [Required] [Column(TypeName = "varchar(50)")] public string FullForm { get; private set; }

        public string StringKey => string.Concat(ShortForm.Take(2));

        public BaseAbbreviation(string shortForm, string fullForm)
        {
            Key = Guid.NewGuid();
            ShortForm = shortForm;
            FullForm = fullForm;
        }

        public string ElementaryRepresentation => GetElementaryRepresentation();
        public string ComplexRepresentation => GetComplexRepresentation();

        public abstract bool IsAbbreviation(string shortForm);
        public abstract string GetFullForm(string shortForm);
        public abstract bool TryGetFullForm(string shortForm, out string fullForm);

        protected abstract string GetElementaryRepresentation();
        protected abstract string GetComplexRepresentation();
    }
}
