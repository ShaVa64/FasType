using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace FasType.Converters.Xaml
{
    public class EnumToColorConverter : IValueConverter
    {
        public object? Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Models.Linguistics.SyllablePosition p)
                return p == Models.Linguistics.SyllablePosition.None ? System.Windows.Media.Brushes.Red : System.Windows.Media.Brushes.Gray;

            return null;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
