using System.Diagnostics;

namespace FasType.Abbreviations
{
    public interface IAbbreviation
    {
        string FullForm { get; }
        string ShortForm { get; }

        bool IsAbbreviation(string shortForm);
        string? GetFullForm(string shortForm);
        bool TryGetFullForm(string shortForm, out string? fullForm);
    }
}