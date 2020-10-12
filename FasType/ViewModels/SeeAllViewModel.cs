using FasType.Models;
using FasType.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;

namespace FasType.ViewModels
{
    public class SeeAllViewModel : BaseViewModel
    {
        readonly IDataStorage _storage;

        public int Count => AllAbbreviations.Count;
        public IList<IAbbreviation> AllAbbreviations { get; private set; }

        public SeeAllViewModel(IDataStorage storage)
        {
            _storage = storage;

            //AllAbbreviations = _storage.ToList();
            AllAbbreviations = _storage.Select(ab => Enumerable.Repeat(ab, 3)).SelectMany(ab => ab).ToList();
        }
    }

    public class SeeAllSelector : DataTemplateSelector
    {
        public DataTemplate First { get; set; }
        public DataTemplate Default { get; set; }
        public DataTemplate Last { get; set; }

        public ItemsControl Control { get; set; }
        public int AlternationCount => Control.AlternationCount;

        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            int alt = ItemsControl.GetAlternationIndex(container);

            if (alt == 0)
                return First;
            if (alt == AlternationCount - 1)
                return Last;
            
            return Default;
            //return base.SelectTemplate(item, container);
        }
    }
}
