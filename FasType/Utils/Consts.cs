using FasType.Core.Models.Abbreviations;
using FasType.Core.Models.Dictionary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FasType.Utils
{
    public static class Abbreviations
    {
        public static readonly BaseAbbreviation OtherAbbreviation = new SimpleAbbreviation("", Properties.Resources.Other, 0, "", "", "");
    }

    public static class DictionaryElements
    {
        public static readonly BaseDictionaryElement OtherElement = new SimpleDictionaryElement(Properties.Resources.Other, "", "", "");
        public static readonly BaseDictionaryElement NoneElement = new SimpleDictionaryElement(Properties.Resources.None, "", "", "");
    }

    public static class Unicodes
    {
        public readonly static string Arrow = "\u2794";
    }
}
