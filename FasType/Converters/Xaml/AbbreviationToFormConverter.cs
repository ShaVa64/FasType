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
            if (values == null || values.Length != 2 || values[0] is not BaseAbbreviation || values[1] is not string)
                return null;

            var a = values[0] as BaseAbbreviation;
            var sf = values[1] as string;
            string ff = a.GetFullForm(sf);

            return string.IsNullOrEmpty(ff) ? a.FullForm : ff;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
