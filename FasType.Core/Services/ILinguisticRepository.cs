using FasType.Core.Contexts;
using FasType.Core.Models.Linguistics;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Core.Services
{
    public interface ILinguisticsRepository : IGenericRepository<AbbreviationMethod, Guid>
    {
        public GrammarType GenderType { get; set; }
        public GrammarType PluralType { get; set; }
        public GrammarType GenderPluralType { get; set; }

        public IEnumerable<AbbreviationMethod> AbbreviationMethods { get; set; }

        string[] Words(string currentWord);
    }

    public class LinguisticsRepository : GenericRepository<AbbreviationMethod, Guid>, ILinguisticsRepository
    {
        public GrammarType GenderType { get => GetGrammarType(); set => SetGrammarType(value); }
        public GrammarType PluralType { get => GetGrammarType(); set => SetGrammarType(value); }
        public GrammarType GenderPluralType { get => GetGrammarType(); set => SetGrammarType(value); }
        public IEnumerable<AbbreviationMethod> AbbreviationMethods { get => GetAll(); set => SetAbbreviationMethods(value); }

        private readonly LinguisticsDbContext _context;

        public LinguisticsRepository(LinguisticsDbContext context) : base(context)
        {
            _context = context;
        }
        void SetAbbreviationMethods(IEnumerable<AbbreviationMethod> enumerable)
        {
            _context.AbbreviationMethods.RemoveRange(AbbreviationMethods);
            _context.AbbreviationMethods.AddRange(enumerable);

            SaveChanges();
        }
        GrammarType GetGrammarType([CallerMemberName] string? name = null)
        {
            _ = name ?? throw new NullReferenceException();
            var gt = _context.GrammarTypes.Find(name);

            return gt ?? new(name, "", GrammarPosition.Prefix);
        }
        void SetGrammarType(GrammarType gt)
        {
            GrammarType record = _context.GrammarTypes.Find(gt.Name);
            if (gt == record)
                return;
            if (record is not null)
                _context.GrammarTypes.Remove(record);
            _context.GrammarTypes.Add(gt);

            SaveChanges();
        }

        const string WC = ".?.?";
        List<string> Words(string _curr, string from, List<string> poss)
        {
            if (from == string.Empty)
            {
                poss.Add(_curr.EndsWith(WC) ? _curr[..^(WC.Length)] : _curr);
                return poss;
            }

            if (_curr != string.Empty && from.Length == 1)
            {
                if ((PluralType.Position == GrammarPosition.Postfix && PluralType.Repr == from)
                    || (GenderType.Position == GrammarPosition.Postfix && GenderType.Repr == from)
                    || (GenderPluralType.Position == GrammarPosition.Postfix && GenderPluralType.Repr == from))
                {
                    poss.Add(_curr);
                }
            }

            AbbreviationMethod[] amrs = Array.Empty<AbbreviationMethod>();
            if (_curr == string.Empty)
            {
                amrs = _context.AbbreviationMethods.Where(m => m.Position.HasFlag(SyllablePosition.Before) && from.StartsWith(m.ShortForm)).ToArray();//.AsEnumerable().Where(m => m.SatisfiesBefore(from)).ToArray();

                if ((PluralType.Position == GrammarPosition.Prefix && PluralType.SuitsGrammar(from))
                    || (GenderType.Position == GrammarPosition.Prefix && GenderType.SuitsGrammar(from))
                    || (GenderPluralType.Position == GrammarPosition.Prefix && GenderPluralType.SuitsGrammar(from)))
                {
                    _ = Words(string.Empty, from[1..], poss);
                }
            }
            else
            {
                amrs = _context.AbbreviationMethods.Where(m => m.Position.HasFlag(SyllablePosition.In) && from.StartsWith(m.ShortForm) && from.Length != m.ShortForm.Length).ToArray();
            }
            amrs = amrs.Concat(_context.AbbreviationMethods.Where(m => m.Position.HasFlag(SyllablePosition.After) && from.EndsWith(m.ShortForm) && from.Length == m.ShortForm.Length)).ToArray();

            _ = Words(_curr + from[0] + WC, from[1..], poss);

            var gAmrs = amrs.GroupBy(amr => amr.ShortForm.Length).ToArray();
            foreach (var g in gAmrs)
            {
                _ = Words(_curr + "(" + string.Join('|', g.Select(amr => amr.FullForm)) + ")" + WC, from[g.Key..], poss);
            }
            return poss;
        }

        public string[] Words(string currentWord) => Words("", currentWord, new()).ToArray();
    }
}
