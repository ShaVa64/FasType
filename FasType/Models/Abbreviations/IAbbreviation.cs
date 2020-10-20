using System.Diagnostics;

namespace FasType.Models.Abbreviations
{
    public interface IAbbreviation
    {
        string FullForm { get; }
        string ShortForm { get; }
        string ElementaryRepresentation { get; }
        string ComplexRepresentation { get; }

        bool IsAbbreviation(string shortForm);
        string GetFullForm(string shortForm);
        bool TryGetFullForm(string shortForm, out string fullForm);
    }
}