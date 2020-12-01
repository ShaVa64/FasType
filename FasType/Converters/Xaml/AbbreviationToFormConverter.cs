using FasType.Models.Abbreviations;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace FasType.Converters.Xaml
{
    public class AbbreviationToFormConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            var a = values[0] as BaseAbbreviation;
            var sf = values[1] as string;
            string ff = a.GetFullForm(sf);

            return ff;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
