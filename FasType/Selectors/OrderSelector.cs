using System.Windows;
using System.Windows.Controls;

namespace FasType.Selectors
{
    public class OrderSelector : DataTemplateSelector
    {
        public DataTemplate First { get; set; }
        public DataTemplate Default { get; set; }
        public DataTemplate Last { get; set; }
        public DataTemplate Only { get; set; }

        public override DataTemplate SelectTemplate(object item, DependencyObject container)
        {
            int altIndex = ItemsControl.GetAlternationIndex(container);

            var ic = ItemsControl.ItemsControlFromItemContainer(container);
            int altCount = ic.AlternationCount;

            if (altCount == 1)
                return Only;

            if (altIndex == 0)
                return First;
            if (altIndex == altCount - 1)
                return Last;
            return Default;
        }
    }
}
