using FasType.Properties;
using FasType.ViewModels;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Resources;
using System.Text;
using System.Windows.Controls;
using System.Windows.Data;

namespace FasType.Converters.Xaml
{
    public class ResourcesToEnumConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture) => throw new NotImplementedException();

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string form = (value as ContentControl).Content as string;

            if (targetType == typeof(FasType.ViewModels.SeeAllViewModel.FormOrderBy))
            {
                if (form == Resources.FullForm)
                    return SeeAllViewModel.FormOrderBy.FullForm;
                else if (form == Resources.ShortForm)
                    return SeeAllViewModel.FormOrderBy.ShortForm;
            }

            return null;
        }
    }
}
