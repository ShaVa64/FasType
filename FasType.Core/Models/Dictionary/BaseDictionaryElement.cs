using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Models.Dictionary
{
    [DebuggerDisplay("{" + nameof(DebuggerDisplay) + "}")]
    public abstract class BaseDictionaryElement
    {
        public string FullForm { get; set; }
        public string[] Others { get; set; }
        public string[] AllForms { get; }
        string DebuggerDisplay => string.Join(", ", AllForms);

        protected BaseDictionaryElement(string fullForm, string[] others)
        {
            FullForm = fullForm;
            Others = others.ToArray();

            AllForms = Others.Prepend(fullForm).ToArray();
        }

        public override bool Equals(object? obj)
        {
            return obj is BaseDictionaryElement element &&
                   FullForm == element.FullForm;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(FullForm);
        }
    }
}
