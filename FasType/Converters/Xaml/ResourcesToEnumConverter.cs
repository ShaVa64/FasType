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
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is SeeAllViewModel.FormOrderBy orderBy)
            {
                return orderBy switch
                {
                    SeeAllViewModel.FormOrderBy.FullForm => Resources.FullForm,
                    SeeAllViewModel.FormOrderBy.ShortForm => Resources.ShortForm,
                    _ => throw new NotImplementedException(),
                };
            }

            throw new NotImplementedException();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string form = (value as ContentControl)?.Content as string ?? throw new NullReferenceException();

            if (targetType == typeof(SeeAllViewModel.FormOrderBy))
            {
                if (form == Resources.FullForm)
                    return SeeAllViewModel.FormOrderBy.FullForm;
                else if (form == Resources.ShortForm)
                    return SeeAllViewModel.FormOrderBy.ShortForm;
            }

            throw new NotImplementedException();
        }
    }
}
