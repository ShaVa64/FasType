using FasType.Models.Linguistics.Grammars;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Windows.Data;

namespace FasType.Converters.Xaml
{
    public class BoolToEnumConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture) => value switch
        {
            GrammarPosition.Postfix => true,
            GrammarPosition.Prefix => false,
            _ => throw new NotImplementedException()
        };

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (targetType == typeof(GrammarPosition))
            {
                return (bool)value switch
                {
                    true => GrammarPosition.Postfix,
                    false => GrammarPosition.Prefix
                };
            }

            throw new NotImplementedException();
        }
    }
}
