using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace FasType.Converters.Xaml
{
    public class IntToStringConverter : IValueConverter
    {
        public enum Parameter
        {
            UsedSeeAll
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var param = (Parameter)parameter;

            return param switch
            {
                Parameter.UsedSeeAll => UsedSeeAll((ulong)value),
                _ => value.ToString()
            };
        }

        static string UsedSeeAll(ulong val) => val == 0 ? "" : $"(x{val})";

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
