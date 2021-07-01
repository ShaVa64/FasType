using FasType.Core.Models.Abbreviations;
using FasType.Core.Services;
using Microsoft.Extensions.DependencyInjection;
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
        private readonly IRepositoriesManager _repositories;
        public AbbreviationToFormConverter()
        {
            _repositories = App.Current.ServiceProvider.GetRequiredService<IRepositoriesManager>();
        }

        public object? Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values == null || values.Length != 2 || values[0] is not BaseAbbreviation || values[1] is not string)
                return null;

            var a = (BaseAbbreviation)values[0];
            var sf = (string)values[1];
            string? ff = a.GetFullForm(sf, _repositories.Linguistics);
            _repositories.Reload();

            return string.IsNullOrEmpty(ff) ? a.FullForm : ff;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture) => throw new NotImplementedException();
    }
}
